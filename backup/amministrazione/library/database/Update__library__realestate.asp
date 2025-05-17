<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-flat
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE NEXT-REALESTATE
'...........................................................................................
function Install__REALESTATE(conn)
	Install__REALESTATE = _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_strutture ("& _
	  "		st_ID " & SQL_PrimaryKey(conn, "Rtb_strutture") + ", "& _
			SQL_MultiLanguageField("st_denominazione_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("st_prezzo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("st_descrizione_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			SQL_MultiLanguageField("st_metratura_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		st_ordine INTEGER NULL, "& _
	  "		st_home BIT NULL, "& _
	  "		st_NEXTweb_ps_mappa_location INTEGER NULL, "& _
	  "		st_NEXTweb_ps_info INTEGER NULL, "& _
	  "		st_NEXTweb_ps_mappa_catastale INTEGER NULL, "& _
	  "		st_tipologia_id INTEGER NULL, "& _
	  "		st_categoria_id INTEGER NULL,"& _
	  "		st_visibile BIT NULL,"& _
	  "		st_indirizzo_mappa " + SQL_CharField(Conn, 255) + " NULL"& _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_tipologie ("& _
			SQL_MultiLanguageField("ti_nome_<lingua>" + SQL_CharField(Conn, 100) + " NULL ") + ", " + _
	  "		ti_ID " & SQL_PrimaryKey(conn, "Rtb_tipologie") + ", "& _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_categorie ("& _
	  "		ca_ID " & SQL_PrimaryKey(conn, "Rtb_categorie") + ", "& _
	  "		car_ordine INTEGER NULL, "& _
	  "		car_icona " + SQL_CharField(Conn, 255) + " NULL, "& _
			SQL_MultiLanguageField("ca_nome_<lingua>" + SQL_CharField(Conn, 100) + " NULL ") + ", " + _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_caratteristiche ("& _
	  "		car_ID " & SQL_PrimaryKey(conn, "Rtb_caratteristiche") + ", "& _
			SQL_MultiLanguageField("car_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		car_ordine INTEGER NULL, "& _
	  "		car_icona " + SQL_CharField(Conn, 255) + " NULL,"& _
	  "		car_tipo INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_strutture_caratteristiche ("& _
	  "		sc_ID " & SQL_PrimaryKey(conn, "Rtb_strutture_caratteristiche") + ", "& _
			SQL_MultiLanguageField("sc_valore_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		sc_struttura_id INTEGER NULL, "& _
	  "		sc_caratteristica_id INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_foto ("& _
	  "		fo_ID " & SQL_PrimaryKey(conn, "Rtb_foto") + ", "& _
	  "		fo_thumb " + SQL_CharField(Conn, 255) + " NULL, "& _
	  "		fo_zoom " + SQL_CharField(Conn, 255) + " NULL, "& _
			SQL_MultiLanguageField("fo_didascalia_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
	  "		fo_ordine INTEGER NULL, "& _
	  "		fo_struttura_id INTEGER NULL,"& _
	  "		fo_visibile BIT NULL,"& _
	  "		fo_numero INT NULL,"& _
	  "		fo_pubblicazione DATETIME NULL"& _
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_richieste_info ("& _
	  "		ri_ID " & SQL_PrimaryKey(conn, "Rtb_richieste_info") + ", "& _
	  "		ri_prezzo MONEY NULL, "& _
	  "		ri_richiesta " + SQL_CharField(Conn, 0) + " NULL, "& _
	  "		ri_NEXTcom_ID INTEGER NULL, "& _
	  "		ri_data DATETIME NULL,"& _
	  "		ri_struttura_id INTEGER NULL"& _
	  ");"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_tipologie "& _
   	  "		FOREIGN KEY (st_tipologia_id) REFERENCES Rtb_tipologie (ti_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_categorie "& _
   	  "		FOREIGN KEY (st_categoria_id) REFERENCES Rtb_categorie (ca_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_strutture_caratteristiche ADD CONSTRAINT FK_Rtb_strutture_caratteristiche__Rtb_strutture "& _
   	  "		FOREIGN KEY (sc_struttura_id) REFERENCES Rtb_strutture (st_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_strutture_caratteristiche ADD CONSTRAINT FK_Rtb_strutture_caratteristiche__Rtb_caratteristiche "& _
   	  "		FOREIGN KEY (sc_caratteristica_id) REFERENCES Rtb_caratteristiche (car_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_foto ADD CONSTRAINT FK_Rtb_foto__Rtb_strutture "& _
   	  "		FOREIGN KEY (fo_struttura_id) REFERENCES Rtb_strutture (st_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_richieste_info ADD CONSTRAINT FK_Rtb_richieste_info__tb_indirizzario "& _
   	  "		FOREIGN KEY (ri_NEXTcom_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Rtb_richieste_info ADD CONSTRAINT FK_Rtb_richieste_info__Rtb_strutture "& _
   	  "		FOREIGN KEY (ri_struttura_id) REFERENCES Rtb_strutture (st_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_citta ("& _
			SQL_MultiLanguageField("ci_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		ci_ID " & SQL_PrimaryKey(conn, "Rtb_citta") + ", "& _
	  ");"& _
	  "ALTER TABLE Rtb_strutture ADD "& _
	  "		st_citta_id INTEGER NULL;"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_citta "& _
   	  "		FOREIGN KEY (st_citta_id) REFERENCES Rtb_citta (ci_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_contratti ("& _
			SQL_MultiLanguageField("co_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		co_ID " & SQL_PrimaryKey(conn, "Rtb_contratti") + ", "& _
	  ");"& _
	  "ALTER TABLE Rtb_strutture ADD "& _
	  "		st_contratto_id INTEGER NULL;"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_contratti "& _
   	  "		FOREIGN KEY (st_contratto_id) REFERENCES Rtb_contratti (co_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE VUOTO
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__VUOTO(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__REALESTATE__1 = ""
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 1
'...........................................................................................
'aggiunge tabelle per gestione listini e fasce di prezzo appartamenti
'...........................................................................................
function Aggiornamento__REALESTATE__1(conn)
	Aggiornamento__REALESTATE__1 = _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "rtb_categorieRealEstate ("& vbCrLf & _
	  "		catC_id " & SQL_PrimaryKey(conn, "rtb_categorieRealEstate") + ", "& vbCrLf & _
			SQL_MultiLanguageField("catC_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		catC_codice " + SQL_CharField(Conn, 50) + " NULL, "& vbCrLf & _
	  "		catC_foglia BIT NULL, "& vbCrLf & _
	  "		catC_livello INTEGER NULL, "& vbCrLf & _
	  "		catC_padre_id INTEGER NULL, "& vbCrLf & _
	  "		catC_ordine INTEGER NULL, "& vbCrLf & _
			SQL_MultiLanguageField("catC_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
	  "		catC_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL"& vbCrLf & _
	  ");"& vbCrLf &_
	  " ALTER TABLE Rtb_strutture DROP Constraint FK_Rtb_strutture__Rtb_categorie;" & vbCRLF & _
	  " ALTER TABLE Rtb_strutture DROP Constraint FK_Rtb_strutture__Rtb_tipologie;" & vbCRLF & _
	  " ALTER TABLE Rtb_strutture DROP COLUMN st_categoria_id;" & vbCRLF &_
	  " ALTER TABLE Rtb_strutture DROP COLUMN st_tipologia_id;" & vbCRLF &_
	  " DROP TABLE Rtb_tipologie; " & vbCRLF & _
	  " DROP TABLE Rtb_categorie; " & vbCRLF & _
 	  " ALTER TABLE Rtb_strutture ADD "& vbCrLf & _
	  "		st_categoria_id INT NULL; "& vbCrLf & _
	  " ALTER TABLE Rtb_strutture ADD "& vbCrLf & _
	  "		st_area_id INT NULL; "& vbCrLf & _
	  " ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__rtb_categorieRealEstate "& vbCrLf & _
   	  "		FOREIGN KEY (st_categoria_id) REFERENCES rtb_categorieRealEstate (catC_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "	CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_descrittori ("& vbCrLf & _
	  "		des_ID " & SQL_PrimaryKey(conn, "Rtb_descrittori") + ", "& vbCrLf & _
			SQL_MultiLanguageField("des_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		des_tipo INTEGER NULL, "& vbCrLf & _
	  "		des_ordine INTEGER NULL, "& vbCrLf & _
	  "		des_principale BIT NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rrel_descrittori_realestate ("& vbCrLf & _
	  "		rdi_ID " & SQL_PrimaryKey(conn, "Rrel_descrittori_realestate") + ", "& vbCrLf & _
	  "		rdi_descrittore_id INTEGER NULL, "& vbCrLf & _
			SQL_MultiLanguageField("rdi_valore_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		rdi_st_id INTEGER NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Rrel_descrittori_realestate ADD CONSTRAINT FK_Rrel_descrittori_realestate__Rtb_descrittori "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_descrittore_id) REFERENCES Rtb_descrittori (des_ID) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE Rrel_descrittori_realestate ADD CONSTRAINT FK_Rrel_descrittori_realestate__Rtb_strutture "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_st_id) REFERENCES Rtb_strutture (st_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 2
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__2(conn)
Aggiornamento__REALESTATE__2 = _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "rtb_Aree ("& vbCrLf & _
	  "		are_id " & SQL_PrimaryKey(conn, "rtb_Aree") + ", "& vbCrLf & _
			SQL_MultiLanguageField("are_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
	  "		are_external_source " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
	  "		are_foto " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
	  "		are_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
	  "		are_codice " + SQL_CharField(Conn, 50) + " NULL, "& vbCrLf & _
	  "		are_external_id " + SQL_CharField(Conn, 50) + " NULL, "& vbCrLf & _
	  "		are_foglia BIT NULL, "& vbCrLf & _
	  "		are_visibile BIT NULL, "& vbCrLf & _
	  "		are_albero_visibile BIT NULL, "& vbCrLf & _
	  "		are_livello INTEGER NULL, "& vbCrLf & _
	  "		are_padre_id INTEGER NULL, "& vbCrLf & _
	  "		are_ordine INTEGER NULL, "& vbCrLf & _
	  "		are_tipologia_padre_base INTEGER NULL, "& vbCrLf & _
			SQL_MultiLanguageField("are_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
	  "		are_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL"& vbCrLf & _
	  ");"& vbCrLf &_
	  " ALTER TABLE Rtb_strutture DROP Constraint FK_Rtb_strutture__Rtb_citta;" & vbCRLF & _
	  " ALTER TABLE Rtb_strutture DROP COLUMN st_citta_id;" & vbCrLf &_
	  " ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__rtb_Aree "& vbCrLf & _
   	  " FOREIGN KEY (st_area_id) REFERENCES rtb_Aree (are_id) "& vbCrLf & _
	  " ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  " DROP TABLE Rtb_citta;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 3
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__3(conn)
	Aggiornamento__REALESTATE__3 = _
	"CREATE TABLE " & SQL_Dbo(Conn) & "Rrel_categorieRealEstate_descrittori (" + vbCrLf + _
	"	rtd_id " & SQL_PrimaryKey(conn, "Rrel_categorieRealEstate_descrittori") + "," + vbCrLf + _
	"	rtd_tipologia_id int NULL ," + vbCrLf + _
	"	rtd_descrittore_id int NULL ," + vbCrLf + _
	"	rtd_ordine int NULL ," + vbCrLf + _
	"	rtd_locked bit NULL " + vbCrLf + _
	" ); " + vbCrLf + _
	" ALTER TABLE Rrel_categorieRealEstate_descrittori ADD CONSTRAINT FK_Rrel_categorieRealEstate_descrittori__rtb_categorieRealEstate "& vbCrLf & _
	" FOREIGN KEY (rtd_tipologia_id) REFERENCES rtb_categorieRealEstate (catC_id) "& vbCrLf & _
	" ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	" ALTER TABLE Rrel_categorieRealEstate_descrittori ADD CONSTRAINT FK_Rrel_categorieRealEstate_descrittori__rtb_Descrittori "& vbCrLf & _
	" FOREIGN KEY (rtd_descrittore_id) REFERENCES rtb_Descrittori (des_id) "& vbCrLf & _
	" ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 4
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__4(conn)
	Aggiornamento__REALESTATE__4 = "ALTER TABLE rtb_categorieRealEstate ADD catC_foto " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 5
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__5(conn)
	Aggiornamento__REALESTATE__5 = "ALTER TABLE rtb_categorieRealEstate ADD catC_albero_visibile BIT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 6
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__6(conn)
	Aggiornamento__REALESTATE__6 = _
		"ALTER TABLE rtb_categorieRealEstate ADD" & _
		"	catC_tipologia_padre_base INTEGER NULL," &_
		" 	catC_visibile BIT NULL," &_
		" 	catC_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 7
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__REALESTATE__7(conn)
	Aggiornamento__REALESTATE__7 = "" &_
		"ALTER TABLE Rtb_strutture_caratteristiche DROP Constraint FK_Rtb_strutture_caratteristiche__Rtb_caratteristiche;" & vbCRLF & _
		"DROP TABLE Rtb_caratteristiche;" & vbCRLF &_
		"DROP TABLE Rtb_strutture_caratteristiche;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 8
'...........................................................................................
'aggiunta relazione tra categorie e utenti dell'area riservata
'...........................................................................................
function Aggiornamento__REALESTATE__8(conn)
	Aggiornamento__REALESTATE__8 = _
		"CREATE TABLE " & SQL_Dbo(Conn) & "Rrel_categorie_utenti (" + vbCrLf + _
		"	rcu_id " & SQL_PrimaryKey(conn, "Rrel_categorie_utenti") + "," + vbCrLf + _
		"	rcu_categoria_id int NULL ," + vbCrLf + _
		"	rcu_utente_id int NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		" ALTER TABLE Rrel_categorie_utenti ADD CONSTRAINT FK_Rtb_categorieRealEstate__Rrel_categorie_utenti"& vbCrLf & _
		" FOREIGN KEY (rcu_categoria_id) REFERENCES Rtb_categorieRealEstate(catC_id) "& vbCrLf & _
		" ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
		" ALTER TABLE Rrel_categorie_utenti ADD CONSTRAINT FK_tb_utenti__Rrel_categorie_utenti"& vbCrLf & _
		" FOREIGN KEY (rcu_utente_id) REFERENCES tb_utenti(ut_id) "& vbCrLf & _
		" ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 9
'...........................................................................................
'aggiunta gestione agenzie e campo intero prezzi
'...........................................................................................
function Aggiornamento__REALESTATE__9(conn)
	Aggiornamento__REALESTATE__9 = _
  	  " CREATE TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ("& vbCrLf & _
	  "		age_id INT NOT NULL, "& vbCrLf & _
	  "		age_admin_id INT NULL, "& vbCrLf & _
	  "		age_gruppo_id INT NULL, "& vbCrLf & _
	  "		age_url " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
	  "		age_url_prenotazione " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
			SQL_MultiLanguageField("age_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
	  "		age_logo " + SQL_CharField(Conn, 255) + " NULL, "& vbCrLf & _
	  "		age_visibile BIT NULL, "& vbCrLf & _
	  "		age_ordine INT NULL "& vbCrLf & _
	  ");"& vbCrLf &_
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ADD CONSTRAINT PK_rtb_agenzie " + vbCrLf + _
	  "		PRIMARY KEY CLUSTERED (age_id);" + vbCrLf + _
	  SQL_AddForeignKey(conn, "rtb_agenzie", "age_id", "tb_indirizzario", "idElencoIndirizzi", true, "") & vbCrLf & _
	  SQL_AddForeignKey(conn, "rtb_agenzie", "age_admin_id", "tb_admin", "id_admin", false, "") & vbCrLf & _
	  SQL_AddForeignKey(conn, "rtb_agenzie", "age_gruppo_id", "tb_gruppi", "id_gruppo", false, "") & vbCrLf & _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_strutture ADD"& vbCrLf & _
	  "		st_prezzoValore_it MONEY NULL, "& vbCrLf & _
	  "		st_prezzoValore_en MONEY NULL, "& vbCrLf & _
	  "		st_prezzoValore_fr MONEY NULL, "& vbCrLf & _
	  "		st_prezzoValore_es MONEY NULL, "& vbCrLf & _
	  "		st_prezzoValore_de MONEY NULL, "& vbCrLf & _
	  "		st_agenzia_id INT NULL" & vbCrLf + _
	  " ; "& vbCrLf & _
	  " INSERT INTO tb_indirizzario(NomeOrganizzazioneElencoIndirizzi, isSocieta, ModoRegistra, DataIscrizione,"& vbCrLf & _
	  " 	LockedByApplication, ApplicationsLocker, lingua)"& vbCrLf & _
	  " 	VALUES ('agenzia NextRealestate', 1, 'agenzia NextRealestate', "& SQL_Now(conn) &", 1, ' "& NEXTREALESTATE &",', 'it');"& vbCrLf & _
	  " INSERT INTO rel_rub_ind (id_indirizzo, id_rubrica)"& vbCrLf & _
	  " 	SELECT TOP 1 idElencoIndirizzi, (SELECT TOP 1 id_rubrica FROM tb_rubriche)"& vbCrLf & _
	  "		FROM tb_indirizzario ORDER BY idElencoIndirizzi DESC;"& vbCrLf & _
	  " INSERT INTO rtb_agenzie (age_id, age_visibile, age_ordine)"& vbCrLf & _
	  " 	SELECT TOP 1 idElencoIndirizzi, 1, 10 FROM tb_indirizzario ORDER BY idElencoIndirizzi DESC;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 10
'...........................................................................................
'termina l'aggiornamento 9 separatamente causa lettura con GetValueList
'...........................................................................................
function Aggiornamento__REALESTATE__10(conn)
	Aggiornamento__REALESTATE__10 = _
  	  " UPDATE rtb_strutture SET st_agenzia_id = "& GetValueList(conn, NULL, "SELECT TOP 1 age_id FROM rtb_agenzie ORDER BY age_id DESC") &";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 11
'...........................................................................................
'termina l'aggiornamento 9 separatamente causa blocco tabelle su access
'...........................................................................................
function Aggiornamento__REALESTATE__11(conn)
	Aggiornamento__REALESTATE__11 = _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_strutture ALTER COLUMN st_agenzia_id INT NOT NULL;"& vbCrLf & _
	  " ALTER TABLE rtb_strutture ADD CONSTRAINT FK_rtb_strutture__rtb_agenzie"& vbCrLf & _
	  " FOREIGN KEY (st_agenzia_id) REFERENCES rtb_agenzie(age_id)"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 12
'...........................................................................................
'aggiunge integrita tra agenzie e utenti dell'area riservata
'...........................................................................................
function Aggiornamento__REALESTATE__12(conn)
	Aggiornamento__REALESTATE__12 = _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ADD st_utente_id INT NULL;"& vbCrLf & _
	  " INSERT INTO tb_utenti (ut_nextCom_id, ut_abilitato)"& vbCrLf & _
	  " 	SELECT TOP 1 age_id, 0 FROM rtb_agenzie ORDER BY age_id DESC;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 13
'...........................................................................................
'termina l'aggiornamento 12 separatamente causa lettura con GetValueList
'...........................................................................................
function Aggiornamento__REALESTATE__13(conn)
	Aggiornamento__REALESTATE__13 = _
	  " UPDATE rtb_agenzie SET st_utente_id = "& GetValueList(conn, NULL, "SELECT TOP 1 ut_id FROM tb_utenti ORDER BY ut_id DESC")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 14
'...........................................................................................
'termina l'aggiornamento 13 separatamente causa blocco tabelle su access
'...........................................................................................
function Aggiornamento__REALESTATE__14(conn)
	Aggiornamento__REALESTATE__14 = _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ALTER COLUMN st_utente_id INT NOT NULL;"& vbCrLf & _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ADD CONSTRAINT FK_rtb_agenzie__tb_utenti"& vbCrLf & _
	  " 	FOREIGN KEY (st_utente_id) REFERENCES tb_utenti(ut_id);"& vbCrLf & _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_strutture ADD"& vbCrLf & _
	  "		st_pub_area_id INT NULL,"& vbCrLf & _
	  "		st_pub_contratto_id INT NULL,"& vbCrLf & _
	  "		st_pub_categoria_id INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 15
'...........................................................................................
'aggiunge l'ID dell'immobile client sul server
'...........................................................................................
function Aggiornamento__REALESTATE__15(conn)
	Aggiornamento__REALESTATE__15 = _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_strutture ADD"& vbCrLf & _
	  "		st_pub_client_id INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 16
'...........................................................................................
'aggiunge l'ID del descrittore server sul client
'...........................................................................................
function Aggiornamento__REALESTATE__16(conn)
	Aggiornamento__REALESTATE__16 = _
	  " ALTER TABLE " & SQL_Dbo(Conn) & "rtb_descrittori ADD"& vbCrLf & _
	  "		des_pub_server_id INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 17
'...........................................................................................
'aggiunge i campi per indicizzare l'appartamento su google maps.
'...........................................................................................
function Aggiornamento__REALESTATE__17(conn)
	Aggiornamento__REALESTATE__17 = _
		"ALTER TABLE Rtb_strutture ADD " + _
		"	st_google_maps_latitudine FLOAT NULL, " + _
		"	st_google_maps_longitudine FLOAT NULL ; " + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 18
'...........................................................................................
'toglie il campo relazione tra agenzie e utenti perche implicito su tb_utenti
'...........................................................................................
function Aggiornamento__REALESTATE__18(conn)
	Aggiornamento__REALESTATE__18 = _
		" ALTER TABLE rtb_agenzie DROP CONSTRAINT FK_rtb_agenzie__tb_utenti;" + vbCrLf + _
		" ALTER TABLE rtb_agenzie DROP COLUMN st_utente_id;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 19
'...........................................................................................
'aggiunge le viste per il recupero e la visualizzazione di immobili ed agenzie
'...........................................................................................
function Aggiornamento__REALESTATE__19(conn)
	Aggiornamento__REALESTATE__19 = _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_agenzie AS " + vbCrLf + _
		"	SELECT *, " + vbCrLf + _
		"		( " & SQL_IF(conn, SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
		"					 ) ", "1", "0") & ") AS age_visibile_assoluto " + vbCrLf + _
		" 	FROM (rtb_agenzie INNER JOIN v_Indirizzario ON rtb_agenzie.age_id = v_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON v_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID " + _
		" ; " + vbCrLF + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_agenzie_visibili AS " + vbCrLf + _
		"	SELECT * " + vbCrLf + _
		" 	FROM (rtb_agenzie INNER JOIN v_Indirizzario ON rtb_agenzie.age_id = v_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON v_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID " + _
		"	WHERE " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"		  ( tb_utenti.ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"									    " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
		" ; " + vbCrLF + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture AS " + vbCrLf + _
		"	SELECT * , " + vbCrLF + _
		"		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
		"					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID) " + vbCrLf + _
		"		INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID " + _
		" ; " + vbCrLF + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili AS " + vbCrLf + _
		"	SELECT * " + vbCrLF + _
		"	FROM (((((Rtb_strutture INNER JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID ) " + vbCrLf + _
		"		INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID " + _
		"	WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
		"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
		"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
		"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
		"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
		"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 20
'...........................................................................................
'aggiunge campo riferimento e relazione tra immobili e contattaci
'...........................................................................................
function Aggiornamento__REALESTATE__20(conn)
	Aggiornamento__REALESTATE__20 = _
		" ALTER TABLE rtb_strutture ADD st_riferimento " + SQL_CharField(Conn, 100) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 21
'...........................................................................................
'corregge problema di installazione next-real estate su database dove non e' attiva l'applicazione
'...........................................................................................
function Aggiornamento__REALESTATE__21(conn)
	Aggiornamento__REALESTATE__21 = _
		" DELETE FROM rtb_agenzie WHERE (SELECT COUNT(*) FROM tb_siti WHERE id_sito=" & NEXTREALESTATE & ")=0 ; " + _
		" DELETE FROM tb_indirizzario WHERE ApplicationsLocker LIKE '%" & NEXTREALESTATE & "%' AND (SELECT COUNT(*) FROM tb_siti WHERE id_sito=" & NEXTREALESTATE & ")=0 ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 22
'...........................................................................................
'aggiunge campi per gestione richiesta su client
'...........................................................................................
function Aggiornamento__REALESTATE__22(conn)
	Aggiornamento__REALESTATE__22 = _
		" ALTER TABLE rtb_richieste_info ADD" + vbCrLF + _
		"	ri_codice " + SQL_CharField(Conn, 255) + "," + vbCrLF + _
		"	ri_pub_id INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 23
'...........................................................................................
'aggiunge campi per mantenimento sincronizzazione con casavenezia.it
'...........................................................................................
function Aggiornamento__REALESTATE__23(conn)
	Aggiornamento__REALESTATE__23 = _
		" ALTER TABLE rtb_strutture ADD" + vbCrLF + _
		"	st_pub_visibile BIT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 24
'...........................................................................................
'aggiunge area alle agenzie. Aggiunge flag di attivazione "scheda completa" dell'agenzia
'...........................................................................................
function Aggiornamento__REALESTATE__24(conn)
	Aggiornamento__REALESTATE__24 = _
		" ALTER TABLE rtb_agenzie ADD" + _
		"	age_area_id INT NULL, " + _
		" 	age_scheda_completa BIT NULL; " + _
		" UPDATE rtb_agenzie SET age_scheda_completa = 1; " + _
		SQL_AddForeignKey(conn, "rtb_agenzie", "age_area_id", "rtb_aree", "are_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 25
'...........................................................................................
'corregge anche viste delle agenzie
'...........................................................................................
function Aggiornamento__REALESTATE__25(conn)
	Aggiornamento__REALESTATE__25 = _
		DropObject(conn, "rv_agenzie", "VIEW") + vbCrLf + _
		DropObject(conn, "rv_agenzie_visibili", "VIEW") + vbCrLf + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_agenzie AS " + vbCrLf + _
		"	SELECT *, " + vbCrLf + _
		"		( " & SQL_IF(conn, SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
		"					 ) ", "1", "0") & ") AS age_visibile_assoluto " + vbCrLf + _
		" 	FROM ((rtb_agenzie INNER JOIN v_Indirizzario ON rtb_agenzie.age_id = v_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON v_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
		"		LEFT JOIN rtb_aree ON rtb_agenzie.age_area_id = rtb_aree.are_id " + _
		" ; " + vbCrLF + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_agenzie_visibili AS " + vbCrLf + _
		"	SELECT * " + vbCrLf + _
		" 	FROM ((rtb_agenzie INNER JOIN v_Indirizzario ON rtb_agenzie.age_id = v_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON v_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
		"		LEFT JOIN rtb_aree ON rtb_agenzie.age_area_id = rtb_aree.are_id " + vbCrLF + _
		"	WHERE " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"		  ( tb_utenti.ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"									    " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 26
'...........................................................................................
'aggiunge campi di log inserimento e modifica per le agenzie e gli immobili
'...........................................................................................
function Aggiornamento__REALESTATE__26(conn)
	Aggiornamento__REALESTATE__26 = _
		" ALTER TABLE rtb_agenzie ADD" + _
			AddInsModFields("age") + "; " + _
		" ALTER TABLE rtb_strutture ADD" + _
			AddInsModFields("st") + "; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__26 = Aggiornamento__REALESTATE__26 + _
			" UPDATE rtb_agenzie " + _
			"	SET age_insData = co_insData, age_insAdmin_id = co_insAdmin_id, age_modData = co_modData, age_modAdmin_id = co_modAdmin_id " + _
			"	FROM ( rtb_agenzie INNER JOIN tb_contents ON tb_contents.co_F_Key_id = rtb_agenzie.age_id ) " + _
			"		 INNER JOIN tb_siti_Tabelle ON (tb_contents.co_F_table_id = tb_siti_Tabelle.tab_id AND tb_siti_Tabelle.tab_name LIKE 'rtb_agenzie') ; " + _
			" UPDATE rtb_strutture " + _
			"	SET st_insData = co_insData, st_insAdmin_id = co_insAdmin_id, st_modData = co_modData, st_modAdmin_id = co_modAdmin_id " + _
			"   FROM ( rtb_strutture INNER JOIN tb_contents ON tb_contents.co_F_Key_id = rtb_strutture.st_id ) " + _
			" 		 INNER JOIN tb_siti_Tabelle ON (tb_contents.co_F_table_id = tb_siti_Tabelle.tab_id AND tb_siti_Tabelle.tab_name LIKE 'rtb_strutture') ; "
	else
		Aggiornamento__REALESTATE__26 = Aggiornamento__REALESTATE__26 + _
			" UPDATE ( rtb_agenzie INNER JOIN tb_contents ON tb_contents.co_F_Key_id = rtb_agenzie.age_id ) " + _
			"		   INNER JOIN tb_siti_Tabelle ON (tb_contents.co_F_table_id = tb_siti_Tabelle.tab_id AND tb_siti_Tabelle.tab_name LIKE 'rtb_agenzie') " + _
			"	SET age_insData = co_insData, age_insAdmin_id = co_insAdmin_id, age_modData = co_modData, age_modAdmin_id = co_modAdmin_id ; " + _
			" UPDATE ( rtb_strutture INNER JOIN tb_contents ON tb_contents.co_F_Key_id = rtb_strutture.st_id ) " + _
			" 		   INNER JOIN tb_siti_Tabelle ON (tb_contents.co_F_table_id = tb_siti_Tabelle.tab_id AND tb_siti_Tabelle.tab_name LIKE 'rtb_strutture') " + _
			"	SET st_insData = co_insData, st_insAdmin_id = co_insAdmin_id, st_modData = co_modData, st_modAdmin_id = co_modAdmin_id ; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 27
'...........................................................................................
'aggiunge campi per mantenimento sincronizzazione con casavenezia.it
'...........................................................................................
function Aggiornamento__REALESTATE__27(conn)
	Aggiornamento__REALESTATE__27 = _
		" SELECT * FROM AA_VERSIONE ; " + _
		AddInsModRelations(conn, "rtb_agenzie", "age") + _
		AddInsModRelations(conn, "rtb_strutture", "st")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 28
'...........................................................................................
'corregge installazione agenzia di default per versioni "mono-agenzia"
'...........................................................................................
function Aggiornamento__REALESTATE__28(conn)
	Aggiornamento__REALESTATE__28 = _
		" UPDATE tb_utenti SET ut_abilitato=1 " + _
		" WHERE ut_nextCom_id IN (SELECT age_id FROM rtb_agenzie) " + _
		"   AND (SELECT COUNT(*) FROM rtb_agenzie)=1 "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 29
'...........................................................................................
'aggiunge campi memo a descrittori degli immobili.
'...........................................................................................
function Aggiornamento__REALESTATE__29(conn)
	Aggiornamento__REALESTATE__29 = _
		" ALTER TABLE Rrel_descrittori_realestate " + SQL_AddColumn(conn) + " " + _
		SQL_MultiLanguageField("rdi_memo_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 30
'...........................................................................................
'aggiunge permesso al NEXT-realestate.
'...........................................................................................
function Aggiornamento__REALESTATE__30(conn)
	Aggiornamento__REALESTATE__30 = _
		" UPDATE tb_siti SET sito_p2 = 'REAL_AGENCY' WHERE sito_dir LIKE '%NextRealEstate%'; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 31
'...........................................................................................
'pulisce l'indirizzo se uguale a 'http://'.
'...........................................................................................
function Aggiornamento__REALESTATE__31(conn)
	Aggiornamento__REALESTATE__31 = _
		" UPDATE rtb_agenzie SET age_url = '' WHERE age_url LIKE 'http://'"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 32
'...........................................................................................
'aggiunge campi descrizione e denominazione alternativi a tabella degli immobili.
'...........................................................................................
function Aggiornamento__REALESTATE__32(conn)
	Aggiornamento__REALESTATE__32 = _
		" ALTER TABLE Rtb_strutture " + SQL_AddColumn(conn) + " " + _
		SQL_MultiLanguageField("st_pub_descrizione_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		SQL_MultiLanguageField("st_pub_denominazione_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 33
'...........................................................................................
'aggiunge raggruppamenti ai descrittori.
'...........................................................................................
function Aggiornamento__REALESTATE__33(conn)
	Aggiornamento__REALESTATE__33 = _
		" ALTER TABLE Rtb_descrittori ADD " + _
		"	des_per_ricerca BIT NULL, " + _
		"	des_raggruppamento_id INT NULL " + _
		" ; " + _
		" CREATE TABLE " & SQL_Dbo(conn) & "Rtb_descrittori_raggruppamenti ( " + vbCrLf + _
		"	desr_id " + SQL_PrimaryKey(conn, "Rtb_descrittori_raggruppamenti") + ", " + _
		SQL_MultiLanguageField("	desr_titolo_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _
		"	desr_ordine int NULL, " + vbCrLf + _
		"	desr_codice " + SQL_CharField(Conn, 255) + ", " + _
		"	desr_di_sistema int NULL" + vbCrLf + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "Rtb_descrittori", "des_raggruppamento_id", "Rtb_descrittori_raggruppamenti", "desr_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 34
'...........................................................................................
'Giacomo 18/12/2009
'aggiunge tabella per tipizzazione delle foto
'...........................................................................................
function Aggiornamento__REALESTATE__34(conn)
	Aggiornamento__REALESTATE__34 = _
		" ALTER TABLE Rtb_foto ADD " + _
		"	fo_tipo_id INT NULL; " + _
		" CREATE TABLE " & SQL_Dbo(conn) & "Rtb_foto_tipo ( " & _
		"	ft_id " + SQL_PrimaryKey(conn, "Rtb_foto_tipo") + ", " + _
		"	ft_nome " + SQL_CharField(Conn, 255) + " NULL, "+ _
		"	ft_codice " + SQL_CharField(Conn, 255) + " NULL " + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "Rtb_foto", "fo_tipo_id", "Rtb_foto_tipo", "ft_id", true, "") + _
		" INSERT INTO Rtb_foto_tipo(ft_nome,ft_codice) VALUES ('immagini', 'img') ; " + _
		" UPDATE Rtb_foto SET fo_tipo_id = 1 ; "
end function
'*******************************************************************************************
	
	
'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 35
'...........................................................................................
'Giacomo 22/12/2009
'
'...........................................................................................
function Aggiornamento__REALESTATE__35(conn)
	Aggiornamento__REALESTATE__35 = _
		" ALTER TABLE Rtb_descrittori ADD " + _
		SQL_MultiLanguageField(" des_unita_<lingua> " + SQL_CharField(Conn, 100)) + "NULL, " + _
		" des_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
		" des_per_confronto BIT NULL, " + _
		" des_img " + SQL_CharField(Conn, 255) + " NULL; "
end function
'*******************************************************************************************



'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************

'AGGIORNAMENTO REALESTATE 36
'...........................................................................................
'	Giacomo, 21/01/2010
'...........................................................................................
' SERIE DI FUNZIONI PER AGGIUNGERE I CAMPI PER UNA NUOVA LINGUA SUL REAL ESTATE
'...........................................................................................

'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************
function Update_language_NextRealEstate(conn, lingua_abbr)
dim sql
	Select case DB_Type(conn)
		case DB_Access
			sql = " ALTER TABLE Rrel_descrittori_realestate ADD " + _
				  " 	rdi_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	rdi_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_agenzie ADD " + _
				  " 	age_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_Aree ADD " + _
				  " 	are_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	are_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_categorieRealEstate ADD " + _
				  " 	catC_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	catC_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE Rtb_contratti ADD " + _
				  " 	co_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
				  " ALTER TABLE Rtb_descrittori ADD " + _
				  " 	des_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	des_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 100) + " NULL;" + _
				  " ALTER TABLE Rtb_descrittori_raggruppamenti ADD " + _
				  " 	desr_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
				  " ALTER TABLE Rtb_foto ADD " + _
				  " 	fo_didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE Rtb_strutture ADD " + _
				  " 	st_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
				  " 	st_denominazione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_prezzo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_metratura_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_prezzoValore_" + lingua_abbr + " MONEY NULL," + _
				  " 	st_pub_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
				  " 	st_pub_denominazione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;"
			if TableExists(conn, "rtb_agenzieCategorie") then
				sql = sql + " 	ALTER TABLE rtb_agenzieCategorie ADD " + _
							" 		ageC_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL, " + _
							" 		ageC_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL; "
			end if
		case DB_SQL
			sql = " ALTER TABLE Rrel_descrittori_realestate ADD " + _
				  " 	rdi_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	rdi_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_agenzie ADD " + _
				  " 	age_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_Aree ADD " + _
				  " 	are_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	are_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE rtb_categorieRealEstate ADD " + _
				  " 	catC_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	catC_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE Rtb_contratti ADD " + _
				  " 	co_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
				  " ALTER TABLE Rtb_descrittori ADD " + _
				  " 	des_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	des_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 100) + " NULL;" + _
				  " ALTER TABLE Rtb_descrittori_raggruppamenti ADD " + _
				  " 	desr_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
				  " ALTER TABLE Rtb_foto ADD " + _
				  " 	fo_didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
				  " ALTER TABLE Rtb_strutture ADD " + _
				  " 	st_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
				  " 	st_denominazione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_prezzo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_metratura_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
				  " 	st_prezzoValore_" + lingua_abbr + " MONEY NULL," + _
				  " 	st_pub_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
				  " 	st_pub_denominazione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;"
			if TableExists(conn, "rtb_agenzieCategorie") then
				  sql = sql + " ALTER TABLE rtb_agenzieCategorie ADD " + _
				  " 			ageC_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL, " + _
				  " 			ageC_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL; "
			end if
		end select
	Update_language_NextRealEstate = sql  
end function


function Aggiornamento__REALESTATE__36(conn, lingua_abbr)
	Aggiornamento__REALESTATE__36 = _
		Update_language_NextRealEstate(conn, lingua_abbr)
end function


function Update_language_NextRealEstate_1(lingua_abbr)
	sql = " ALTER TABLE rtb_agenzie ADD " + _
		  "		age_marchio_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
	Update_language_NextRealEstate_1 = Update_language_NextRealEstate(conn, lingua_abbr) + sql
end function
	
	



'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************

function Update_language_Cancella_lingua(lingua_abbr)
	Update_language_Cancella_lingua = " ALTER TABLE Rrel_descrittori_realestate DROP COLUMN " + _
		  " 	rdi_valore_" + lingua_abbr + " ," + _
		  " 	rdi_memo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE rtb_agenzie DROP COLUMN " + _
		  " 	age_descr_" + lingua_abbr + " ;" + _
		  " ALTER TABLE rtb_Aree DROP COLUMN " + _
		  " 	are_nome_" + lingua_abbr + " ," + _
		  " 	are_descr_" + lingua_abbr + " ;" + _
		  " ALTER TABLE rtb_categorieRealEstate DROP COLUMN " + _
		  " 	catC_nome_" + lingua_abbr + " ," + _
		  " 	catC_descr_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Rtb_contratti DROP COLUMN " + _
		  " 	co_nome_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Rtb_descrittori DROP COLUMN " + _
		  " 	des_nome_" + lingua_abbr + " ," + _
		  " 	des_unita_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Rtb_descrittori_raggruppamenti DROP COLUMN " + _
		  " 	desr_titolo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Rtb_foto DROP COLUMN " + _
		  " 	fo_didascalia_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Rtb_strutture DROP COLUMN " + _
		  " 	st_descrizione_" + lingua_abbr + " ," + _
		  " 	st_denominazione_" + lingua_abbr + " ," + _
		  " 	st_prezzo_" + lingua_abbr + " ," + _
		  " 	st_metratura_" + lingua_abbr + " ," + _
		  " 	st_prezzoValore_" + lingua_abbr + " ," + _
		  " 	st_pub_descrizione_" + lingua_abbr + " ," + _
		  " 	st_pub_denominazione_" + lingua_abbr + " ;" 

end function

function AggiornamentoSpeciale__REALESTATE__36(conn, lingua_abbr)
	AggiornamentoSpeciale__REALESTATE__36 = _
		Update_language_Cancella_lingua(lingua_abbr)
end function

'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 37
'...........................................................................................
'Giacomo 15/03/2010
'Aggiunge tabella delle categorie per le agenzie
'...........................................................................................
function Aggiornamento__REALESTATE__37(conn)
	Aggiornamento__REALESTATE__37 = _
		"CREATE TABLE " & SQL_Dbo(Conn) & "rtb_agenzieCategorie (" & _
		"		ageC_id " & SQL_PrimaryKey(conn, "rtb_agenzieCategorie") + ", " & _
				SQL_MultiLanguageFieldComplete(conn, "ageC_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " & _
		"		ageC_codice " + SQL_CharField(Conn, 50) + " NULL, " & _
		"		ageC_foglia BIT NULL, " & _
		"		ageC_livello INT NULL, " & _
		"		ageC_padre_id INT NULL, " & _
		"		ageC_ordine INT NULL, " & _
				SQL_MultiLanguageFieldComplete(conn, "ageC_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " & _
		"		ageC_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL, " & _
		"		ageC_foto " + SQL_CharField(Conn, 255) + " NULL, " & _
		"		ageC_albero_visibile BIT NULL, " & _
		"		ageC_tipologia_padre_base INT NULL, " & _
		" 		ageC_visibile BIT NULL, " & _
		" 		ageC_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL " & _
		");" & _
		" ALTER TABLE rtb_agenzie ADD " & _
		" 	age_categoria_id INT NULL; " & _
		SQL_AddForeignKey(conn, "rtb_agenzie", "age_categoria_id", "rtb_agenzieCategorie", "ageC_id", true, "") & _
		" INSERT INTO rtb_agenzieCategorie(ageC_nome_it, ageC_nome_en, ageC_tipologia_padre_base, ageC_tipologie_padre_lista, ageC_livello, ageC_foglia ) VALUES ('agenzia', 'agency', 1, 1, 0, 1) ; " & _
		" UPDATE rtb_agenzie SET age_categoria_id = 1 ; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 38
'...........................................................................................
'Giacomo 17/03/2010
'modifica a rtb_strutture per gestione condomini
'...........................................................................................
function Aggiornamento__REALESTATE__38(conn)
	Aggiornamento__REALESTATE__38 = _
		" ALTER TABLE Rtb_strutture ADD " & _
		" 	st_is_condominio BIT NULL, " & _
		" 	st_condominio_id INT NULL; " & _
		SQL_AddForeignKey(conn, "Rtb_strutture", "st_condominio_id", "Rtb_strutture", "st_ID", false, "") & _
		SQL_RemoveForeignKey(conn, "Rtb_strutture", "st_contratto_id", "Rtb_contratti", true, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 39
'...........................................................................................
'Giacomo 23/03/2010
'cancello e ricreo rv_strutture perchè non è più obbligatorio il contratto su rtb_strutture
'...........................................................................................
function Aggiornamento__REALESTATE__39(conn)
	Aggiornamento__REALESTATE__39 = _
		" DROP VIEW rv_strutture; " + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture AS " + vbCrLf + _
		"	SELECT * , " + vbCrLF + _
		"		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
		"					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		" ; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 40
'...........................................................................................
'Giacomo 23/03/2010
'aggiunge paramentro nell'area amministrativa per abilitare o disabilitare i condomini
'...........................................................................................
function Aggiornamento__REALESTATE__40(conn)
	Aggiornamento__REALESTATE__40 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__40(conn)
	if GetValueList(conn, NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
		CALL AddParametroSito(conn, "ABILITA_CONDOMINI", _
									null, _
									"Abilita la gestione dei condomini", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__40 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 41
'...........................................................................................
'Giacomo 24/03/2010
'cancello e ricreo rv_strutture_visibili perchè non è più obbligatorio il contratto su rtb_strutture
'...........................................................................................
function Aggiornamento__REALESTATE__41(conn)
	Aggiornamento__REALESTATE__41 = _
		" DROP VIEW rv_strutture_visibili; " + _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili AS " + vbCrLf + _
				"	SELECT * " + vbCrLF + _
				"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + _
				"	WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 42
'...........................................................................................
'Giacomo 25/03/2010
'aggiunta parametri copiandoli da parametri old
'...........................................................................................
function Aggiornamento__REALESTATE__42(conn)
	Aggiornamento__REALESTATE__42 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__42(conn)
	if GetValueList(conn, NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
		CALL AddParametroSito(conn, "AGENZIE_ABILITATE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'AGENZIE_ABILITATE' AND par_sito_id=18"), null, null, null, null)
		
		CALL AddParametroSito(conn, "AREA_RISERVATA_AGENZIE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'AREA_RISERVATA_AGENZIE' AND par_sito_id=18"), null, null, null, null)
		
		CALL AddParametroSito(conn, "CASAVENEZIA_PRINCIPALE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_PRINCIPALE' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "CASAVENEZIA_STANDALONE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_STANDALONE' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "CASAVENEZIA_WSDL", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_WSDL' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "CASAVENEZIA_WSDL_LOGIN", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_WSDL_LOGIN' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "CASAVENEZIA_WSDL_RUBRICA_RICHIESTE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_WSDL_RUBRICA_RICHIESTE' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "CATEGORIE_STRUTTURE_ABILITATE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CATEGORIE_ABILITATE' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "PERMESSI_AREARISERVATA_ABILITATI", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'PERMESSI_AREARISERVATA_ABILITATI' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "PREZZI_NUMERICI_ABILITATI", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'PREZZI_NUMERICI_ABILITATI' AND par_sito_id=18"), null, null, null, null)
								
		CALL AddParametroSito(conn, "RUBRICA_AGENZIE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'RUBRICA_AGENZIE' AND par_sito_id=18"), null, null, null, null)

		CALL AddParametroSito(conn, "RUBRICA_AGENZIE_AFFILIATE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'RUBRICA_AGENZIE_AFFILIATE' AND par_sito_id=18"), null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__42 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************
	

'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 43
'...........................................................................................
'Giacomo 26/03/2010
'creo la vista per i condomini rv_condomini
'...........................................................................................
function Aggiornamento__REALESTATE__43(conn)
	Aggiornamento__REALESTATE__43 = _
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini AS " + vbCrLf + _
		"	SELECT * , " + vbCrLF + _
		"		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
		"					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
		"					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
		"										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
		"					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"	WHERE " & SQL_IsTrue(conn, "st_is_condominio") & ";"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 44
'...........................................................................................
'Giacomo 29/03/2010
'aggiunge paramentro nell'area amministrativa per abilitare l'amministrazione delle AREE alle agenzie
'...........................................................................................
function Aggiornamento__REALESTATE__44(conn)
	Aggiornamento__REALESTATE__44 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__44(conn)
	if GetValueList(conn, NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
		CALL AddParametroSito(conn, "ABILITA_AREE_PER_AGENZIE", _
									null, _
									"Abilita la gestione delle aree alle agenzie", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 45
'...........................................................................................
'Giacomo 15/04/2010
'modifica a rtb_agenzie 
'...........................................................................................
function Aggiornamento__REALESTATE__45(conn)
	Aggiornamento__REALESTATE__45 = _
		" ALTER TABLE rtb_agenzie ADD " & _
				SQL_MultiLanguageFieldComplete(conn, "age_marchio_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + "; " 
		Select case DB_Type(conn)
			case DB_SQL
				Aggiornamento__REALESTATE__45 = Aggiornamento__REALESTATE__45 + _
					" UPDATE rtb_agenzie SET age_marchio_it = (SELECT NomeOrganizzazioneElencoIndirizzi FROM tb_Indirizzario  WHERE tb_Indirizzario.IDElencoIndirizzi = rtb_agenzie.age_id) ; "
		end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 46
'...........................................................................................
'Giacomo 19/04/2010
'aggiunta parametri copiandoli da parametri old - CORREZIONE Aggiornamento 42
'...........................................................................................
function Aggiornamento__REALESTATE__46(conn)
	Aggiornamento__REALESTATE__46 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__46(conn)
	if GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
		CALL AddParametroSito(conn, "CASAVENEZIA_WSDL_PASSWORD", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_WSDL_PASSWORD' AND par_sito_id=18"), null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__46 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 47
'...........................................................................................
'Nicola 03/05/2010
'aggiunge campo interno ad amministrazione a gestione immobili
'...........................................................................................
function Aggiornamento__REALESTATE__47(conn)
	Aggiornamento__REALESTATE__47 = _
		" ALTER TABLE rtb_strutture ADD " & _
		"	st_proprietario " + SQL_CharField(Conn, 255) + " NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 48
'...........................................................................................
'Andrea 04/05/2010
'crea le viste suddivise per lingua per rv_condomini
'...........................................................................................
function Aggiornamento__REALESTATE__48(conn)

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = " SELECT   Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLf + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLf + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLf + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "	
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_cn AS " + vbCrLf + _
		" SELECT  Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_descrizione_cn, " + vbCrLf + _
                      "Rtb_strutture.st_denominazione_cn, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_metratura_cn, Rtb_strutture.st_prezzoValore_cn, " + vbCrLf + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_nome_cn, rtb_Aree.are_descr_cn, rtb_categorieRealEstate.catC_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_it, rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_descr_cn, " + vbCrLf + _
                      "rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, " + vbCrLf + _
                      "rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, " + vbCrLf + _
                      "rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, " + vbCrLf + _
                      "rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, Rtb_contratti.co_nome_cn,  " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__48 = _
		DropObject(conn,"rv_condomini_it","VIEW") + _
		DropObject(conn,"rv_condomini_en","VIEW") + _
		DropObject(conn,"rv_condomini_fr","VIEW") + _
		DropObject(conn,"rv_condomini_de","VIEW") + _
		DropObject(conn,"rv_condomini_es","VIEW") + _
		DropObject(conn,"rv_condomini_pt","VIEW") + _
		DropObject(conn,"rv_condomini_ru","VIEW") + _
		DropObject(conn,"rv_condomini_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__48 = Aggiornamento__REALESTATE__48 + Agg_ru + Agg_cn + Agg_pt
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 49
'...........................................................................................
'Andrea 04/05/2010
'crea le viste suddivise per lingua per rv_strutture
'...........................................................................................
function Aggiornamento__REALESTATE__49(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, " + vbCrLF + _
                      "Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _ 
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__49 = _
		DropObject(conn,"rv_strutture_it","VIEW") + _
		DropObject(conn,"rv_strutture_en","VIEW") + _
		DropObject(conn,"rv_strutture_fr","VIEW") + _
		DropObject(conn,"rv_strutture_de","VIEW") + _
		DropObject(conn,"rv_strutture_es","VIEW") + _
		DropObject(conn,"rv_strutture_ru","VIEW") + _
		DropObject(conn,"rv_strutture_pt","VIEW") + _
		DropObject(conn,"rv_strutture_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__49 = Aggiornamento__REALESTATE__49 + Agg_ru + Agg_cn + Agg_pt
	end if	
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 50
'...........................................................................................
'Andrea 04/05/2010
'crea le viste suddivise per lingua per rv_strutture_visibili
'...........................................................................................
function Aggiornamento__REALESTATE__50(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID " + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, " + vbCrLF + _
                      "Rtb_strutture.st_proprietario, rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID" + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__50 = _
		DropObject(conn,"rv_strutture_visibili_it","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_en","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_fr","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_de","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_es","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_ru","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_pt","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__50 = Aggiornamento__REALESTATE__50 + Agg_ru + Agg_cn + Agg_pt
	end if		
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 51
'...........................................................................................
'Giacomo 12/05/2010
'aggiunta parametri 
'...........................................................................................
function Aggiornamento__REALESTATE__51(conn)
	Aggiornamento__REALESTATE__51 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__51(conn)
	if GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
			CALL AddParametroSito(conn, "CONDOMINI_INSERIMENTO_RAPIDO", _
										null, _
										"Condomini: abilita inserimento rapido immobili", _
										"", _
										adBoolean, _
										0, _
										"", _
										1, _
										1, _
										NEXTREALESTATE, _
										0, null, null, null, null)
			CALL AddParametroSito(conn, "CONDOMINI_IMPORT_IMMOBILI", _
										null, _
										"Condomini: abilita sezione per l'associazione immobili da import dati", _
										"", _
										adBoolean, _
										0, _
										"", _
										1, _
										1, _
										NEXTREALESTATE, _
										0, null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__51 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 52
'...........................................................................................
'Giacomo 12/05/2010
'aggiunta parametro per attivare o disattivare la relazione tra agenzie e descrittori
'...........................................................................................
function Aggiornamento__REALESTATE__52(conn)
	Aggiornamento__REALESTATE__52 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__52(conn)
	if GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
			CALL AddParametroSito(conn, "RELAZIONE_AGENZIE_DESCRITTORI", _
										null, _
										"Attiva la possibilità di scegliere i descrittori per ogni singola agenzia", _
										"", _
										adBoolean, _
										0, _
										"", _
										1, _
										1, _
										NEXTREALESTATE, _
										0, null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__52 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 53
'...........................................................................................
'Giacomo 25/05/2010
'Aggiunge tabella di relazione tra agenzie e descrittori
'...........................................................................................
function Aggiornamento__REALESTATE__53(conn)
	Aggiornamento__REALESTATE__53 = _
		"CREATE TABLE " & SQL_Dbo(Conn) & "Rrel_agenzie_descrittori (" & _
		"		rad_id " & SQL_PrimaryKey(conn, "Rrel_agenzie_descrittori") + ", " & _
		"		rad_agenzia_id INT NULL, " & _
		"		rad_descrittore_id INT NULL " & _
		");" & _
		SQL_AddForeignKey(conn, "Rrel_agenzie_descrittori", "rad_agenzia_id", "rtb_agenzie", "age_id", true, "") & _
		SQL_AddForeignKey(conn, "Rrel_agenzie_descrittori", "rad_descrittore_id", "Rtb_descrittori", "des_ID", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 54
'...........................................................................................
'Matteo 25/05/2010
'aggiunta valore numerico superficie in rtb_strutture
'...........................................................................................
function Aggiornamento__REALESTATE__54(conn)
	Aggiornamento__REALESTATE__54 = _
		" ALTER TABLE " & SQL_Dbo(Conn) & "Rtb_strutture ADD " & _
		SQL_MultiLanguageFieldComplete(conn, "st_metraturaValore_<lingua> INT NULL ") + "; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 55
'...............................................................................................
'Matteo 26/05/2010
'crea le viste suddivise per lingua per rv_condomini (aggiunta campi superficie valore numerico)
'...............................................................................................
function Aggiornamento__REALESTATE__55(conn)

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = " SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLf + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLf + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLf + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "	
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_cn AS " + vbCrLf + _
		" SELECT       Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_descrizione_cn, " + vbCrLf + _
                      "Rtb_strutture.st_denominazione_cn, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_metratura_cn, Rtb_strutture.st_prezzoValore_cn, " + vbCrLf + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLf + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_nome_cn, rtb_Aree.are_descr_cn, rtb_categorieRealEstate.catC_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_it, rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_descr_cn, " + vbCrLf + _
                      "rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, " + vbCrLf + _
                      "rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, " + vbCrLf + _
                      "rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, " + vbCrLf + _
                      "rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, Rtb_contratti.co_nome_cn,  " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__55 = _
		DropObject(conn,"rv_condomini_it","VIEW") + _
		DropObject(conn,"rv_condomini_en","VIEW") + _
		DropObject(conn,"rv_condomini_fr","VIEW") + _
		DropObject(conn,"rv_condomini_de","VIEW") + _
		DropObject(conn,"rv_condomini_es","VIEW") + _
		DropObject(conn,"rv_condomini_pt","VIEW") + _
		DropObject(conn,"rv_condomini_ru","VIEW") + _
		DropObject(conn,"rv_condomini_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__55 = Aggiornamento__REALESTATE__55 + Agg_ru + Agg_cn + Agg_pt
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 56
'...............................................................................................
'Matteo 26/05/2010
'crea le viste suddivise per lingua per rv_strutture (aggiunta campi superficie valore numerico)
'...............................................................................................
function Aggiornamento__REALESTATE__56(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _ 
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__56 = _
		DropObject(conn,"rv_strutture_it","VIEW") + _
		DropObject(conn,"rv_strutture_en","VIEW") + _
		DropObject(conn,"rv_strutture_fr","VIEW") + _
		DropObject(conn,"rv_strutture_de","VIEW") + _
		DropObject(conn,"rv_strutture_es","VIEW") + _
		DropObject(conn,"rv_strutture_ru","VIEW") + _
		DropObject(conn,"rv_strutture_pt","VIEW") + _
		DropObject(conn,"rv_strutture_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__56 = Aggiornamento__REALESTATE__56 + Agg_ru + Agg_cn + Agg_pt
	end if	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 57
'........................................................................................................
'Matteo 26/05/2010
'crea le viste suddivise per lingua per rv_strutture_visibili (aggiunta campi superficie valore numerico)
'........................................................................................................
function Aggiornamento__REALESTATE__57(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID " + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID" + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__57 = _
		DropObject(conn,"rv_strutture_visibili_it","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_en","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_fr","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_de","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_es","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_ru","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_pt","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__57 = Aggiornamento__REALESTATE__57 + Agg_ru + Agg_cn + Agg_pt
	end if		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 58
'...........................................................................................
'Giacomo 01/06/2010
'aggiunta campo url in rtb_strutture
'...........................................................................................
function Aggiornamento__REALESTATE__58(conn)
	Aggiornamento__REALESTATE__58 = _
		" ALTER TABLE " & SQL_Dbo(Conn) & "Rtb_strutture ADD "
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__REALESTATE__58 = Aggiornamento__REALESTATE__58 + _
				SQL_MultiLanguageFieldComplete(conn, "st_url_<lingua> " + SQL_CharField(Conn, 500) + " NULL ") + "; "
		else
			Aggiornamento__REALESTATE__58 = Aggiornamento__REALESTATE__58 + _
				SQL_MultiLanguageFieldComplete(conn, "st_url_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "; "
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 59
'...........................................................................................
'Giacomo 01/06/2010
'aggiunta campo url in rtb_agenzie
'...........................................................................................
function Aggiornamento__REALESTATE__59(conn)
	Aggiornamento__REALESTATE__59 = _
		" ALTER TABLE " & SQL_Dbo(Conn) & "rtb_agenzie ADD "
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__REALESTATE__59 = Aggiornamento__REALESTATE__59 + _
				SQL_MultiLanguageFieldComplete(conn, "age_url_<lingua> " + SQL_CharField(Conn, 500) + " NULL ") + "; "
		else
			Aggiornamento__REALESTATE__59 = Aggiornamento__REALESTATE__59 + _
				SQL_MultiLanguageFieldComplete(conn, "age_url_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "; "
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 60
'...............................................................................................
'Giacomo 01/06/2010
'crea le viste suddivise per lingua per rv_condomini (aggiunta campi url su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__60(conn)

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = " SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLf + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLf + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLf + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "	
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_cn AS " + vbCrLf + _
		" SELECT       Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_descrizione_cn, " + vbCrLf + _
                      "Rtb_strutture.st_denominazione_cn, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_metratura_cn, Rtb_strutture.st_prezzoValore_cn, " + vbCrLf + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLf + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_nome_cn, rtb_Aree.are_descr_cn, rtb_categorieRealEstate.catC_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_it, rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_descr_cn, " + vbCrLf + _
                      "rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, " + vbCrLf + _
                      "rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, " + vbCrLf + _
                      "rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, " + vbCrLf + _
                      "rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, Rtb_contratti.co_nome_cn,  " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__60 = _
		DropObject(conn,"rv_condomini_it","VIEW") + _
		DropObject(conn,"rv_condomini_en","VIEW") + _
		DropObject(conn,"rv_condomini_fr","VIEW") + _
		DropObject(conn,"rv_condomini_de","VIEW") + _
		DropObject(conn,"rv_condomini_es","VIEW") + _
		DropObject(conn,"rv_condomini_pt","VIEW") + _
		DropObject(conn,"rv_condomini_ru","VIEW") + _
		DropObject(conn,"rv_condomini_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__60 = Aggiornamento__REALESTATE__60 + Agg_ru + Agg_cn + Agg_pt
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 61
'...............................................................................................
'Giacomo 01/06/2010
'crea le viste suddivise per lingua per rv_strutture (aggiunta campi url su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__61(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _ 
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__61 = _
		DropObject(conn,"rv_strutture_it","VIEW") + _
		DropObject(conn,"rv_strutture_en","VIEW") + _
		DropObject(conn,"rv_strutture_fr","VIEW") + _
		DropObject(conn,"rv_strutture_de","VIEW") + _
		DropObject(conn,"rv_strutture_es","VIEW") + _
		DropObject(conn,"rv_strutture_ru","VIEW") + _
		DropObject(conn,"rv_strutture_pt","VIEW") + _
		DropObject(conn,"rv_strutture_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__61 = Aggiornamento__REALESTATE__61 + Agg_ru + Agg_cn + Agg_pt
	end if	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 62
'........................................................................................................
'Giacomo 01/06/2010
'crea le viste suddivise per lingua per rv_strutture_visibili (aggiunta campi url su rtb_strutture)
'........................................................................................................
function Aggiornamento__REALESTATE__62(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID " + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID" + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__62 = _
		DropObject(conn,"rv_strutture_visibili_it","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_en","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_fr","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_de","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_es","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_ru","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_pt","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__62 = Aggiornamento__REALESTATE__62 + Agg_ru + Agg_cn + Agg_pt
	end if		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 63
'...........................................................................................
'Nicola 03/06/2010
'aggiunta indici su tabella strutture e revisione indice della primary key per ordinare
'in modo nativo per id decrescente
'...........................................................................................
function Aggiornamento__REALESTATE__63(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__63 = _
			"CREATE NONCLUSTERED INDEX [IDX_rtb_strutture_categoria] ON [dbo].[Rtb_strutture] " + _
			"	(st_categoria_id ASC); " + _
			"CREATE NONCLUSTERED INDEX [IDX_rtb_strutture_area] ON [dbo].[Rtb_strutture] " + _
			"	(st_area_id ASC); " + _
			"CREATE NONCLUSTERED INDEX [IDX_rtb_strutture_agenzia] ON [dbo].[Rtb_strutture] " + _ 
			"	(st_agenzia_id ASC); " + _
			SQL_RemoveForeignKey(conn, "Rtb_foto", "fo_struttura_id", "Rtb_strutture", true, "") + _
			SQL_RemoveForeignKey(conn, "Rtb_richieste_info", "ri_struttura_id", "Rtb_strutture", true, "") + _
			SQL_RemoveForeignKey(conn, "Rrel_descrittori_realestate", "rdi_st_id", "Rtb_strutture", true, "") + _
			SQL_RemoveForeignKey(conn, "Rtb_strutture", "st_condominio_id", "Rtb_strutture", false, "FK_Rtb_strutture__st_condominio_id") + _
			" ALTER TABLE Rtb_strutture DROP CONSTRAINT PK_Rtb_strutture ; " + _
			" ALTER TABLE Rtb_strutture ADD  CONSTRAINT PK_Rtb_strutture " + _
			"	PRIMARY KEY CLUSTERED (st_ID DESC) ; " + _
			SQL_AddForeignKeyExtended(conn, "Rtb_foto", "fo_struttura_id", "Rtb_strutture", "st_id", true, true, "") + _
			SQL_AddForeignKeyExtended(conn, "Rtb_richieste_info", "ri_struttura_id", "Rtb_strutture", "st_id", true, true, "") + _
			SQL_AddForeignKeyExtended(conn, "Rrel_descrittori_realestate", "rdi_st_id", "Rtb_strutture", "st_id", true, true, "") + _
			SQL_AddForeignKey(conn, "Rtb_strutture", "st_condominio_id", "Rtb_strutture", "st_ID", false, "") + _
			"CREATE NONCLUSTERED INDEX [IX_Rtb_foto] ON [dbo].[Rtb_foto] " + _
			"( fo_struttura_id ASC ); "			
	else
		Aggiornamento__REALESTATE__63 = "SELECT * FROM AA_versione"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 64
'...........................................................................................
'Giacomo 12/05/2010
'aggiunta parametro per attivare o disattivare la relazione tra agenzie e descrittori
'...........................................................................................
function Aggiornamento__REALESTATE__64(conn)
	Aggiornamento__REALESTATE__64 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__64(conn)
	if GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = 18") <> "" then
			CALL AddParametroSito(conn, "CASAVENEZIA_AGENZIE_ABILITATE", _
									null, _
									"(PARAMETRI OLD)", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									GetValueList(conn, NULL, "SELECT par_value FROM tb_siti_parametri WHERE par_key LIKE 'CASAVENEZIA_AGENZIE_ABILITATE' AND par_sito_id=18"), null, null, null, null)

	end if
	AggiornamentoSpeciale__REALESTATE__64 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 65
'...........................................................................................
'Nicola 22/11/2010
'aggiunta indici su tabella descrittori
'...........................................................................................
function Aggiornamento__REALESTATE__65(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__65 = _
			" CREATE NONCLUSTERED INDEX [idx_Rrel_descrittori_realestate] ON Rrel_descrittori_realestate (" + vbCrLf + _
			" 	rdi_descrittore_id ASC, " + vbCrLF + _
			"	rdi_st_id ASC " + vbCrLf + _
			" ) ; "		
	else
		Aggiornamento__REALESTATE__65 = "SELECT * FROM AA_versione"
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 66
'...........................................................................................
'Giacomo 16/12/2010
'aggiunta parametro decidere se gestire automaticamente l'import foto per immobili (creazione automatica di una cartella per immobile)
'...........................................................................................
function Aggiornamento__REALESTATE__66(conn)
	Aggiornamento__REALESTATE__66 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__66(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTREALESTATE)) <> "" then
			CALL AddParametroSito(conn, "GESTIONE_CARTELLE_IMMOBILI", _
									null, _
									"crea e gestisci automaticamente una cartella per ogni immobile per l'import foto", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)

	end if
	AggiornamentoSpeciale__REALESTATE__66 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 67
'...........................................................................................
'	Giacomo, 19/05/2011
'...........................................................................................
'   crea indici per ottimizzazione indice e pagine
'...........................................................................................
function Aggiornamento__REALESTATE__67(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__REALESTATE__67 = _
				" CREATE INDEX [IDX_rtb_strutture_contratti] ON [dbo].[Rtb_strutture] " + vbcRLF + _
				" ( " + vbcRLF + _
				" 	  [st_contratto_id] ASC " + vbcRLF + _
				" ); " + vbcRLF + _
				DropObject(conn, "[Rtb_strutture].[IDX_rtb_strutture_area]", "INDEX")
		case DB_Access
			Aggiornamento__REALESTATE__67 = _
				" CREATE INDEX [IDX_rtb_strutture_contratti] ON [Rtb_strutture] " + vbcRLF + _
				" ( " + vbcRLF + _
				" 	  [st_contratto_id] ASC " + vbcRLF + _
				" ); " + vbcRLF + _
				DropObject(conn, "[IDX_rtb_strutture_area]", "INDEX")
	end select		
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 68
'...........................................................................................
'	Giacomo, 20/05/2011
'...........................................................................................
'   aggiunta colonna
'...........................................................................................
function Aggiornamento__REALESTATE__68(conn)
	Aggiornamento__REALESTATE__68 = _
		" ALTER TABLE Rtb_strutture ADD " & _
			" 	st_foto_thumb " + SQL_CharField(Conn, 255) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 69
'...........................................................................................
'	Giacomo, 20/05/2011
'	aggiunta campo url in rtb_aree
'...........................................................................................
function Aggiornamento__REALESTATE__69(conn)
	Aggiornamento__REALESTATE__69 = _
		" ALTER TABLE " & SQL_Dbo(Conn) & "Rtb_aree ADD "
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__REALESTATE__69 = Aggiornamento__REALESTATE__69 + _
				SQL_MultiLanguageFieldComplete(conn, "are_url_<lingua> " + SQL_CharField(Conn, 500) + " NULL ") + "; "
		else
			Aggiornamento__REALESTATE__69 = Aggiornamento__REALESTATE__69 + _
				SQL_MultiLanguageFieldComplete(conn, "are_url_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "; "
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 70
'...............................................................................................
'	Giacomo, 20/05/2011
'	crea le viste suddivise per lingua per rv_condomini (aggiunta campi url su Rtb_aree e campo thumb su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__70(conn)

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = " SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLf + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLf + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLf + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "	
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_cn AS " + vbCrLf + _
		" SELECT       Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_descrizione_cn, " + vbCrLf + _
                      "Rtb_strutture.st_denominazione_cn, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_metratura_cn, Rtb_strutture.st_prezzoValore_cn, " + vbCrLf + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLf + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_nome_cn, rtb_Aree.are_descr_cn, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_it, rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_descr_cn, " + vbCrLf + _
                      "rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, " + vbCrLf + _
                      "rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, " + vbCrLf + _
                      "rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, " + vbCrLf + _
                      "rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, Rtb_contratti.co_nome_cn,  " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__70 = _
		DropObject(conn,"rv_condomini_it","VIEW") + _
		DropObject(conn,"rv_condomini_en","VIEW") + _
		DropObject(conn,"rv_condomini_fr","VIEW") + _
		DropObject(conn,"rv_condomini_de","VIEW") + _
		DropObject(conn,"rv_condomini_es","VIEW") + _
		DropObject(conn,"rv_condomini_pt","VIEW") + _
		DropObject(conn,"rv_condomini_ru","VIEW") + _
		DropObject(conn,"rv_condomini_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__70 = Aggiornamento__REALESTATE__70 + Agg_ru + Agg_cn + Agg_pt
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 71
'...............................................................................................
'	Giacomo, 20/05/2011
'	crea le viste suddivise per lingua per rv_strutture (aggiunta campi url su Rtb_aree e campo thumb su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__71(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _ 
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__71 = _
		DropObject(conn,"rv_strutture_it","VIEW") + _
		DropObject(conn,"rv_strutture_en","VIEW") + _
		DropObject(conn,"rv_strutture_fr","VIEW") + _
		DropObject(conn,"rv_strutture_de","VIEW") + _
		DropObject(conn,"rv_strutture_es","VIEW") + _
		DropObject(conn,"rv_strutture_ru","VIEW") + _
		DropObject(conn,"rv_strutture_pt","VIEW") + _
		DropObject(conn,"rv_strutture_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__71 = Aggiornamento__REALESTATE__71 + Agg_ru + Agg_cn + Agg_pt
	end if	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 72
'........................................................................................................
'	Giacomo, 20/05/2011
'	crea le viste suddivise per lingua per rv_strutture_visibili (aggiunta campi url su Rtb_aree e campo thumb su rtb_strutture)
'........................................................................................................
function Aggiornamento__REALESTATE__72(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID " + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, " + vbCrLF + _
					  "tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID" + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__72 = _
		DropObject(conn,"rv_strutture_visibili_it","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_en","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_fr","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_de","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_es","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_ru","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_pt","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__72 = Aggiornamento__REALESTATE__72 + Agg_ru + Agg_cn + Agg_pt
	end if		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 73
'...........................................................................................
'Giacomo 25/01/2013
'aggiunge parametro per nascondere sezione "PAGINE COLLEGATE" nella scheda immobile e parametro per attivare CKEditor nel NextRealEstate
'...........................................................................................
function Aggiornamento__REALESTATE__73(conn)
	Aggiornamento__REALESTATE__73 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__73(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTREALESTATE)) <> "" then
			CALL AddParametroSito(conn, "REALESTATE_NASCONDI_SEZIONE_PAGINE_COLLEGATE", _
									null, _
									"Se selezionato, nasconde la sezione 'PAGINE COLLEGATE' dalla scheda immobile.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
									
			CALL AddParametroSito(conn, "REALESTATE_ATTIVA_CKEditor", _
									null, _
									"Attiva CKEditor nel Next-RealEstate", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
			
			if cIntero(GetValueList(conn , NULL, "SELECT sid_id FROM tb_siti_descrittori WHERE sid_codice LIKE 'GESTIONE_CARTELLE_IMMOBILI_PATH'")) = 0 then
				CALL AddParametroSito(conn, "GESTIONE_CARTELLE_IMMOBILI_PATH", _
										null, _
										"percorso - parte finale - per la gestione automatica delle cartelle per le foto degli immobili", _
										"", _
										adVarChar, _
										0, _
										"", _
										1, _
										1, _
										NEXTREALESTATE, _
										null, null, null, null, null)
			end if

	end if
	AggiornamentoSpeciale__REALESTATE__73 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 74
'...........................................................................................
'Giacomo 21/10/2013
'Aggiunge tabella di relazione tra agenzie e categorie
'...........................................................................................
function Aggiornamento__REALESTATE__74(conn)
	Aggiornamento__REALESTATE__74 = _
		"CREATE TABLE " & SQL_Dbo(Conn) & "Rrel_agenzie_categorieRealEstate (" & _
		"		rac_id " & SQL_PrimaryKey(conn, "Rrel_agenzie_categorieRealEstate") + ", " & _
		"		rac_agenzia_id INT NULL, " & _
		"		rac_categoria_id INT NULL " & _
		");" & _
		SQL_AddForeignKey(conn, "Rrel_agenzie_categorieRealEstate", "rac_agenzia_id", "rtb_agenzie", "age_id", true, "") & _
		SQL_AddForeignKey(conn, "Rrel_agenzie_categorieRealEstate", "rac_categoria_id", "rtb_categorieRealEstate", "catC_id", true, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 75
'...........................................................................................
'	Giacomo, 27/02/2014
'...........................................................................................
'   aggiunta colonne per dati riservati su tabella immobili
'...........................................................................................
function Aggiornamento__REALESTATE__75(conn)
	Aggiornamento__REALESTATE__75 = _
		" ALTER TABLE Rtb_strutture ADD " + _
		" 	st_is_riservato bit NULL, " + _
		SQL_MultiLanguageFieldComplete(conn, "st_prezzo_riservato_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	st_indirizzo_riservato " + SQL_CharField(Conn, 255) + " NULL, " + _
		"	st_google_maps_latitudine_r FLOAT NULL, " + _
		"	st_google_maps_longitudine_r FLOAT NULL, " + _
		SQL_MultiLanguageFieldComplete(conn, "st_descrizione_riservata_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		SQL_MultiLanguageFieldComplete(conn, "st_note_riservate_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		"	st_nome_proprietario " + SQL_CharField(Conn, 255) + " NULL, " + _
		"	st_dati_proprietario " + SQL_CharField(Conn, 0) + " NULL; " + _
		" ALTER TABLE Rtb_foto ADD " + _
		" 	fo_is_riservata bit NULL; "
end function

function AggiornamentoSpeciale__REALESTATE__75(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTREALESTATE)) <> "" then
			CALL AddParametroSito(conn, "REALESTATE_ATTIVA_DATI_RISERVATI_IMMOBILI", _
									null, _
									"Attiva la gestione dei dati riservati per gli immobili", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__75 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 76
'...............................................................................................
'	Giacomo, 27/02/2014
'	crea le viste suddivise per lingua per rv_condomini (aggiunta colonne dati riservati su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__76(conn)

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = " SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, " + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _					  
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLf + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLf + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLf + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		" WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "	
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_condomini_cn AS " + vbCrLf + _
		" SELECT       Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLf + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLf + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLf + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLf + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLf + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLf + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLf + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLf + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLf + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_descrizione_cn, " + vbCrLf + _
                      "Rtb_strutture.st_denominazione_cn, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_metratura_cn, Rtb_strutture.st_prezzoValore_cn, " + vbCrLf + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, " + vbCrLf + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLf + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, Rtb_strutture.st_prezzo_riservato_cn," + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, Rtb_strutture.st_descrizione_riservata_cn, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, Rtb_strutture.st_note_riservate_cn, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLf + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLf + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLf + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLf + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_nome_cn, rtb_Aree.are_descr_cn, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_nome_it, rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLf + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_descr_cn, " + vbCrLf + _
                      "rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, " + vbCrLf + _
                      "rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, " + vbCrLf + _
                      "rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, " + vbCrLf + _
                      "rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, " + vbCrLf + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLf + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLf + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLf + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLf + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLf + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLf + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLf + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLf + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, Rtb_contratti.co_nome_cn,  " + vbCrLf + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "     "	& SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "			( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "     "	& SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
		"	FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
		"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
		"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
		"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
		"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLf + _
		"WHERE  " & SQL_IsTrue(conn, "Rtb_strutture.st_is_condominio") & "; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__76 = _
		DropObject(conn,"rv_condomini_it","VIEW") + _
		DropObject(conn,"rv_condomini_en","VIEW") + _
		DropObject(conn,"rv_condomini_fr","VIEW") + _
		DropObject(conn,"rv_condomini_de","VIEW") + _
		DropObject(conn,"rv_condomini_es","VIEW") + _
		DropObject(conn,"rv_condomini_pt","VIEW") + _
		DropObject(conn,"rv_condomini_ru","VIEW") + _
		DropObject(conn,"rv_condomini_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__76 = Aggiornamento__REALESTATE__76 + Agg_ru + Agg_cn + Agg_pt
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 77
'...............................................................................................
'	Giacomo, 27/02/2014
'	crea le viste suddivise per lingua per rv_strutture (aggiunta colonne dati riservati su rtb_strutture)
'...............................................................................................
function Aggiornamento__REALESTATE__77(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, " + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, Rtb_strutture.st_prezzo_riservato_cn," + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, Rtb_strutture.st_descrizione_riservata_cn, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, Rtb_strutture.st_note_riservate_cn, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _ 
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID, " + vbCrLF + _
					  "		( " & SQL_IF(conn, SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
					  "					" & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
					  "					 ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
					  "										 " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) " + vbCrLF + _
					  "					 ) ", "1", "0") & ") AS st_visibile_assoluto " + vbCrLf + _
				"FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLf + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__77 = _
		DropObject(conn,"rv_strutture_it","VIEW") + _
		DropObject(conn,"rv_strutture_en","VIEW") + _
		DropObject(conn,"rv_strutture_fr","VIEW") + _
		DropObject(conn,"rv_strutture_de","VIEW") + _
		DropObject(conn,"rv_strutture_es","VIEW") + _
		DropObject(conn,"rv_strutture_ru","VIEW") + _
		DropObject(conn,"rv_strutture_pt","VIEW") + _
		DropObject(conn,"rv_strutture_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__77 = Aggiornamento__REALESTATE__77 + Agg_ru + Agg_cn + Agg_pt
	end if	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 78
'........................................................................................................
'	Giacomo, 27/02/2014
'	crea le viste suddivise per lingua per rv_strutture_visibili (aggiunta colonne dati riservati su rtb_strutture)
'........................................................................................................
function Aggiornamento__REALESTATE__78(conn)
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_prezzo_it, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_en, Rtb_strutture.st_descrizione_it, Rtb_strutture.st_descrizione_en, Rtb_strutture.st_metratura_it, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_en, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_agenzia_id, Rtb_strutture.st_pub_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, Rtb_strutture.st_is_condominio, " + vbCrLF + _
                      "Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, " + vbCrLF + _
                      "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, " + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _
					  "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_codice, rtb_categorieRealEstate.catC_foglia, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, rtb_categorieRealEstate.catC_ordine, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, rtb_categorieRealEstate.catC_ordine_assoluto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foto, rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_visibile, rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, " + vbCrLF + _
                      "rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, " + vbCrLF + _
                      "rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, " + vbCrLF + _
                      "rtb_agenzie.age_area_id, rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, " + vbCrLF + _
                      "rtb_agenzie.age_modData, rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_en, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_ID " + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "rv_strutture_visibili_cn AS " + vbCrLf + _
		"SELECT     Rtb_strutture.st_ID, Rtb_strutture.st_denominazione_it, Rtb_strutture.st_denominazione_en, Rtb_strutture.st_denominazione_cn, " + vbCrLF + _
                      "Rtb_strutture.st_prezzo_it, Rtb_strutture.st_prezzo_en, Rtb_strutture.st_prezzo_cn, Rtb_strutture.st_descrizione_it, " + vbCrLF + _
                      "Rtb_strutture.st_descrizione_en, Rtb_strutture.st_descrizione_cn, Rtb_strutture.st_metratura_it, Rtb_strutture.st_metratura_en, " + vbCrLF + _
                      "Rtb_strutture.st_metratura_cn, Rtb_strutture.st_ordine, Rtb_strutture.st_home, Rtb_strutture.st_NEXTweb_ps_mappa_location, " + vbCrLF + _
                      "Rtb_strutture.st_NEXTweb_ps_info, Rtb_strutture.st_NEXTweb_ps_mappa_catastale, Rtb_strutture.st_visibile, " + vbCrLF + _
                      "Rtb_strutture.st_indirizzo_mappa, Rtb_strutture.st_contratto_id, Rtb_strutture.st_categoria_id, Rtb_strutture.st_area_id, " + vbCrLF + _
                      "Rtb_strutture.st_prezzoValore_it, Rtb_strutture.st_prezzoValore_en, Rtb_strutture.st_prezzoValore_cn, Rtb_strutture.st_agenzia_id, " + vbCrLF + _
                      "Rtb_strutture.st_pub_area_id, Rtb_strutture.st_pub_contratto_id, Rtb_strutture.st_pub_categoria_id, Rtb_strutture.st_pub_client_id, " + vbCrLF + _
                      "Rtb_strutture.st_google_maps_latitudine, Rtb_strutture.st_google_maps_longitudine, Rtb_strutture.st_riferimento, " + vbCrLF + _
                      "Rtb_strutture.st_pub_visibile, Rtb_strutture.st_insData, Rtb_strutture.st_insAdmin_id, Rtb_strutture.st_modData, " + vbCrLF + _
                      "Rtb_strutture.st_modAdmin_id, Rtb_strutture.st_pub_descrizione_it, Rtb_strutture.st_pub_descrizione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_descrizione_cn, Rtb_strutture.st_pub_denominazione_it, Rtb_strutture.st_pub_denominazione_en, " + vbCrLF + _
                      "Rtb_strutture.st_pub_denominazione_cn, Rtb_strutture.st_is_condominio, Rtb_strutture.st_condominio_id, Rtb_strutture.st_proprietario, " + vbCrLF + _
                      "Rtb_strutture.st_metraturaValore_it, Rtb_strutture.st_metraturaValore_en, Rtb_strutture.st_metraturaValore_cn, " + vbCrLf + _
					  "Rtb_strutture.st_url_it, Rtb_strutture.st_url_en, Rtb_strutture.st_url_cn, Rtb_strutture.st_foto_thumb, " + vbCrLf + _
					  "Rtb_strutture.st_is_riservato, Rtb_strutture.st_prezzo_riservato_it, Rtb_strutture.st_prezzo_riservato_en, Rtb_strutture.st_prezzo_riservato_cn," + vbCrLf + _
					  "Rtb_strutture.st_indirizzo_riservato, Rtb_strutture.st_google_maps_latitudine_r, Rtb_strutture.st_google_maps_longitudine_r, " + vbCrLf + _
					  "Rtb_strutture.st_descrizione_riservata_it, Rtb_strutture.st_descrizione_riservata_en, Rtb_strutture.st_descrizione_riservata_cn, " + vbCrLf + _
					  "Rtb_strutture.st_note_riservate_it, Rtb_strutture.st_note_riservate_en, Rtb_strutture.st_note_riservate_cn, " + vbCrLf + _
					  "Rtb_strutture.st_nome_proprietario, Rtb_strutture.st_dati_proprietario, " + vbCrLf + _
                      "rtb_Aree.are_id, rtb_Aree.are_nome_it, rtb_Aree.are_nome_en, rtb_Aree.are_nome_cn, " + vbCrLF + _
                      "rtb_Aree.are_external_source, rtb_Aree.are_foto, rtb_Aree.are_tipologie_padre_lista, rtb_Aree.are_codice, " + vbCrLF + _
                      "rtb_Aree.are_external_id, rtb_Aree.are_foglia, rtb_Aree.are_visibile, rtb_Aree.are_albero_visibile, rtb_Aree.are_livello, " + vbCrLF + _
                      "rtb_Aree.are_padre_id, rtb_Aree.are_ordine, rtb_Aree.are_tipologia_padre_base, rtb_Aree.are_descr_it, rtb_Aree.are_descr_en, " + vbCrLF + _
                      "rtb_Aree.are_descr_cn, rtb_Aree.are_ordine_assoluto, rtb_Aree.are_url_it, rtb_Aree.are_url_en, rtb_Aree.are_url_cn, rtb_categorieRealEstate.catC_id, rtb_categorieRealEstate.catC_nome_it, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_nome_en, rtb_categorieRealEstate.catC_nome_cn, rtb_categorieRealEstate.catC_codice, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_foglia, rtb_categorieRealEstate.catC_livello, rtb_categorieRealEstate.catC_padre_id, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_ordine, rtb_categorieRealEstate.catC_descr_it, rtb_categorieRealEstate.catC_descr_en, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_descr_cn, rtb_categorieRealEstate.catC_ordine_assoluto, rtb_categorieRealEstate.catC_foto, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_albero_visibile, rtb_categorieRealEstate.catC_tipologia_padre_base, rtb_categorieRealEstate.catC_visibile, " + vbCrLF + _
                      "rtb_categorieRealEstate.catC_tipologie_padre_lista, rtb_agenzie.age_id, rtb_agenzie.age_admin_id, rtb_agenzie.age_gruppo_id, " + vbCrLF + _
                      "rtb_agenzie.age_url, rtb_agenzie.age_url_prenotazione, rtb_agenzie.age_descr_it, rtb_agenzie.age_descr_en, " + vbCrLF + _
                      "rtb_agenzie.age_descr_cn, rtb_agenzie.age_logo, rtb_agenzie.age_visibile, rtb_agenzie.age_ordine, rtb_agenzie.age_area_id, " + vbCrLF + _
                      "rtb_agenzie.age_scheda_completa, rtb_agenzie.age_insData, rtb_agenzie.age_insAdmin_id, rtb_agenzie.age_modData, " + vbCrLF + _
                      "rtb_agenzie.age_modAdmin_id, rtb_agenzie.age_categoria_id, rtb_agenzie.age_marchio_it, rtb_agenzie.age_marchio_en, " + vbCrLF + _
                      "rtb_agenzie.age_marchio_cn, rtb_agenzie.age_url_it, rtb_agenzie.age_url_en, rtb_agenzie.age_url_cn, tb_Indirizzario.IDElencoIndirizzi, tb_Indirizzario.NomeElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.SecondoNomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.QualificaElencoIndirizzi, " + vbCrLF + _
					  "tb_Indirizzario.IndirizzoElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CittaElencoIndirizzi, tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, tb_Indirizzario.DTNASCElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.DataIscrizione, tb_Indirizzario.LockedByApplication, " + vbCrLF + _
                      "tb_Indirizzario.ApplicationsLocker, tb_Indirizzario.SyncroKey, tb_Indirizzario.SyncroTable, tb_Indirizzario.SyncroApplication, " + vbCrLF + _
                      "tb_Indirizzario.LocalitaElencoIndirizzi, tb_Indirizzario.PraticaPrefisso, tb_Indirizzario.PraticaCount, tb_Indirizzario.LuogoNascita, " + vbCrLF + _
                      "tb_Indirizzario.CF, tb_Indirizzario.cntRel, tb_Indirizzario.lingua, tb_Indirizzario.NoteElencoIndirizzi, " + vbCrLF + _
                      "tb_Indirizzario.codiceInserimento, tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine, " + vbCrLF + _
                      "tb_Indirizzario.CntSede, tb_Indirizzario.partita_iva, tb_Utenti.ut_ID, tb_Utenti.ut_NextCom_ID, tb_Utenti.ut_login, " + vbCrLF + _
                      "tb_Utenti.ut_password, tb_Utenti.ut_Abilitato, tb_Utenti.ut_ScadenzaAccesso, Rtb_contratti.co_nome_it, " + vbCrLF + _
                      "Rtb_contratti.co_nome_en, Rtb_contratti.co_nome_cn, Rtb_contratti.co_ID" + vbCrLF + _
		" FROM (((((Rtb_strutture INNER JOIN rtb_Aree ON Rtb_strutture.st_area_id = rtb_Aree.are_id ) " + vbCrLF + _
				"		INNER JOIN rtb_categorieRealEstate ON Rtb_strutture.st_categoria_id = rtb_categorieRealEstate.catC_id ) " + vbCrLF + _
				"		INNER JOIN rtb_agenzie ON Rtb_strutture.st_agenzia_id = rtb_agenzie.age_id ) " + vbCrLF + _
				"		INNER JOIN tb_Indirizzario ON rtb_agenzie.age_id = tb_Indirizzario.IDElencoIndirizzi ) " + vbCrLF + _
				"		LEFT JOIN tb_Utenti ON tb_Indirizzario.IDElencoIndirizzi = tb_Utenti.ut_NextCom_ID) " + vbCrLF + _
				"		LEFT JOIN Rtb_contratti ON Rtb_strutture.st_contratto_id = Rtb_contratti.co_ID " + vbCrLF + _
		" WHERE " & SQL_IsTrue(conn, "st_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "catC_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "are_albero_visibile") & " AND " + vbCrLF + _
				"		  " & SQL_IsTrue(conn, "age_visibile") & " AND " + vbCrLf + _
				"		  ( ut_ID IS NULL OR (" & SQL_IsTrue(conn, "ut_Abilitato") & " AND " + vbCrLF + _
				"			 				  " & SQL_IfIsNull(conn, "ut_ScadenzaAccesso", "(" & SQL_Now(conn) & " + 1)") & " > (" & SQL_now(conn) & " - 1)) )" + vbCrLF + _
				" ; "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__REALESTATE__78 = _
		DropObject(conn,"rv_strutture_visibili_it","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_en","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_fr","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_de","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_es","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_ru","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_pt","VIEW") + _
		DropObject(conn,"rv_strutture_visibili_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__REALESTATE__78 = Aggiornamento__REALESTATE__78 + Agg_ru + Agg_cn + Agg_pt
	end if		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 79
'...........................................................................................
'Giacomo 31/03/2014
'aggiunge paramentro nell'area amministrativa
'...........................................................................................
function Aggiornamento__REALESTATE__79(conn)
	Aggiornamento__REALESTATE__79 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__REALESTATE__79(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTREALESTATE)) <> "" then
			CALL AddParametroSito(conn, "REALESTATE_ATTIVA_INSERIMENTO_RAPIDO_DATI_IMMOBILE", _
									null, _
									"Attiva la gestione per l'inserimento rapido dei dati degli immobili", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTREALESTATE, _
									null, null, null, null, null)
	end if
	AggiornamentoSpeciale__REALESTATE__79 = " SELECT * FROM AA_Versione "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 80
'...........................................................................................
'Giacomo 29/07/2014
'aggiunge paramentro nell'area amministrativa
'...........................................................................................
function Aggiornamento__REALESTATE__80(conn)
	Aggiornamento__REALESTATE__80 = "" & _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_strutture_log ("& _
	  "		sl_id " & SQL_PrimaryKey(conn, "Rtb_strutture_log") & ", "& _
	  "		sl_struttura_nome " & SQL_CharField(Conn, 255) & "NULL, " & _
	  "		sl_url " & SQL_CharField(Conn, 255) & "NULL, " & _
	  "		sl_struttura_id int NULL, "&_
	  "		sl_agenzia_id int NULL, "&_
	  "		sl_user_id int NULL, "&_
	  "		sl_http_request " & SQL_CharField(Conn, 0) & " NULL, "& _
	  "		sl_datetime smalldatetime NULL "&_
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_ricerche_log ("& _
	  "		rl_id " & SQL_PrimaryKey(conn, "Rtb_ricerche_log") & ", "& _
	  "		rl_url " & SQL_CharField(Conn, 255) & "NULL, " & _
	  "		rl_user_id int NULL, "&_
	  "		rl_agenzia_id int NULL, "&_
	  "		rl_http_request " & SQL_CharField(Conn, 0) & " NULL, "& _
	  "		rl_datetime smalldatetime "&_
	  ");"& _
	  "CREATE TABLE " & SQL_Dbo(Conn) & "Rtb_ricerche_log_dettagli ("& _
	  "		rld_id " & SQL_PrimaryKey(conn, "Rtb_ricerche_log_dettagli") & ", "& _
	  "		rld_ricerca_id int NULL, "&_
	  "		rld_descrizione " & SQL_CharField(Conn, 500) & " NULL, "& _
	  "		rld_filtro " & SQL_CharField(Conn, 500) & " NULL, "& _
	  "		rld_tabella " & SQL_CharField(Conn, 255) & " NULL "& _
	  ");" & _
	  SQL_AddForeignKey(conn, "Rtb_ricerche_log_dettagli", "rld_ricerca_id", "Rtb_ricerche_log", "rl_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO REALESTATE 81
'...........................................................................................
'Nicola 24/04/2015
'aggiunta indici su tabella strutture per velocizzazione import dati
'...........................................................................................
function Aggiornamento__REALESTATE__81(conn)
		Aggiornamento__REALESTATE__81 = _
			"CREATE NONCLUSTERED INDEX [IDX_rtb_strutture_pub_client_id] ON [dbo].[Rtb_strutture] " + _
			"	(st_pub_client_id ASC); "		
end function
'*******************************************************************************************


%>

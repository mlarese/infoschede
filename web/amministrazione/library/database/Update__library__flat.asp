<%
'...........................................................................................
'...........................................................................................
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-flat
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE NEXT-FLAT
'...........................................................................................
function Install__FLAT(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FLAT = ""
		case DB_SQL
			Install__FLAT = _
			  "CREATE TABLE Atb_Prenotazioni (" &_
			  "		pre_ID INT IDENTITY (1, 1) NOT NULL, " &_
			  "		pre_NextCom_ID INT NOT NULL, " &_
			  "		pre_ap_ID INT NOT NULL, " &_
			  " 	pre_dataArrivo SMALLDATETIME NULL, " &_
			  " 	pre_NumeroNotti INT NULL, " &_
			  " 	pre_prezzo REAL NULL, " &_
			  "		pre_confermata BIT, " & _
			  "		pre_posti INT NULL, " & _
			  "		CONSTRAINT PK_Atb_Prenotazioni PRIMARY KEY (pre_id) " & _
			  ");" &_
			  "CREATE TABLE Atb_Appartamenti (" &_
			  "		ap_ID INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_Atb_Appartamenti PRIMARY KEY, " &_
			  "		ap_denominazione varchar(250) NULL ," &_
			  "		ap_descrizione_IT text NULL," &_
			  "		ap_descrizione_EN text NULL," &_
			  "		ap_descrizione_FR text NULL," &_
			  "		ap_descrizione_DE text NULL," &_
			  "		ap_descrizione_ES text NULL," &_
			  "		ap_NextWeb_ps_mappa INT NULL, " &_
			  "		ap_NextWeb_ps_gallery INT NULL, " &_
			  "		ap_foto varchar(250) NULL ," & _
			  "		ap_minimo_notti INT NULL, " & _
			  "		ap_posti_letto INT NULL ," & _
			  "		ap_home BIT NULL ," & _
			  "		ap_notti_1 INT NULL, " & _
			  "		ap_sconto_1 REAL NULL, " & _
			  "		ap_notti_2 INT NULL, " & _
			  "		ap_sconto_2 REAL NULL, " & _
			  "		ap_notti_3 INT NULL, " & _
			  "		ap_sconto_3 REAL NULL, " & _
			  "		ap_notti_4 INT NULL, " & _
			  "		ap_sconto_4 REAL NULL, " & _
			  "		ap_notti_5 INT NULL, " & _
			  "		ap_sconto_5 REAL NULL, " & _
			  "		ap_ordine INT NULL, " & _
			  "		ap_abilitato BIT NULL " & _
			  ");" & _
			  "CREATE TABLE Atb_ap_disponibilita (" &_
			  "		dispo_ID INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_Atb_ap_disponibilita PRIMARY KEY, " &_
			  "		dispo_ap_ID INT NOT NULL, " &_
			  "		dispo_data SMALLDATETIME NULL, " &_
			  "		dispo_disponibile BIT NULL ," &_
			  "		dispo_prezzo REAL NOT NULL " &_
			  ");" &_
			  "CREATE TABLE Atb_ap_dotazioni (" &_
			  "		dot_ID INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_Atb_ap_dotazioni PRIMARY KEY, " &_
			  "		dot_ap_ID INT NOT NULL, " &_
			  "		dot_ordine char(1) NULL ," &_
			  "		dot_valore_IT nvarchar(250) NULL," &_
			  "		dot_valore_EN nvarchar(250) NULL," &_
			  "		dot_valore_FR nvarchar(250) NULL," &_
			  "		dot_valore_DE nvarchar(250) NULL," &_
			  "		dot_valore_ES nvarchar(250) NULL" &_
			  ");" &_
			  " ALTER TABLE Atb_Prenotazioni ADD CONSTRAINT FK_Atb_prenotazioni__Atb_Appartamenti " &_
		   	  " 	FOREIGN KEY (pre_ap_id) REFERENCES Atb_Appartamenti (ap_ID) " &_
			  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
			  " ALTER TABLE Atb_Prenotazioni ADD CONSTRAINT FK_Atb_prenotazioni__tb_Indirizzario " &_
		   	  " 	FOREIGN KEY (pre_NextCom_ID) REFERENCES Tb_Indirizzario (IDElencoIndirizzi) " &_
			  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
			  " ALTER TABLE Atb_ap_disponibilita ADD CONSTRAINT FK_Atb_ap_disponibilita__Atb_Appartamenti " &_
		   	  " 	FOREIGN KEY (dispo_ap_ID) REFERENCES Atb_Appartamenti (ap_ID) " &_
			  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
			  " ALTER TABLE Atb_ap_dotazioni ADD CONSTRAINT FK_Atb_ap_dotazioni__Atb_Appartamenti " &_
		   	  " 	FOREIGN KEY (dot_ap_ID) REFERENCES Atb_Appartamenti (ap_ID) " &_
			  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
			  "CREATE INDEX IDX_Atb_Prenotazioni__pre_NextCom_ID ON Atb_Prenotazioni (pre_NextCom_ID);" &_
			  "CREATE INDEX IDX_Atb_Prenotazioni__pre_ap_ID ON Atb_Prenotazioni (pre_ap_ID);" &_
			  "CREATE INDEX IDX_Atb_ap_disponibilita__dispo_ap_ID ON Atb_ap_disponibilita (dispo_ap_ID);" &_
			  "CREATE INDEX IDX_Atb_ap_dotazioni__dot_ap_ID ON Atb_ap_dotazioni (dot_ap_ID)"
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FLAT 1
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__FLAT__1(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FLAT__1 = "ALTER TABLE atb_ap_dotazioni ALTER COLUMN"+ _
									 "	dot_ordine INTEGER NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FLAT 2
'...........................................................................................
'aggiunge tabelle per gestione listini e fasce di prezzo appartamenti
'...........................................................................................
function Aggiornamento__FLAT__2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FLAT__2 = _
				"CREATE TABLE Atb_listini (" + _
	  			"	lis_id COUNTER CONSTRAINT PK_atb_listini PRIMARY KEY ," + _
	  			"	lis_nome_it TEXT(250) WITH COMPRESSION NULL ," + _
				"	lis_nome_en TEXT(250) WITH COMPRESSION NULL ," + _
				"	lis_nome_fr TEXT(250) WITH COMPRESSION NULL ," + _
				"	lis_nome_de TEXT(250) WITH COMPRESSION NULL ," + _
				"	lis_nome_es TEXT(250) WITH COMPRESSION NULL ," + _
				"	lis_data DATETIME NULL, " + _
				"	lis_condizioni_it TEXT WITH COMPRESSION NULL, " + _
	  			"	lis_condizioni_en TEXT WITH COMPRESSION NULL, " + _
				"	lis_condizioni_fr TEXT WITH COMPRESSION NULL, " + _
				"	lis_condizioni_es TEXT WITH COMPRESSION NULL, " + _
				"	lis_condizioni_de TEXT WITH COMPRESSION NULL " + _
			  	" ) ; " + _
				" CREATE TABLE Atb_listini_fasce ( " + _
				"	fas_id COUNTER CONSTRAINT PK_atb_listini_fasce PRIMARY KEY , " + _
				"	fas_nome_it TEXT(250) WITH COMPRESSION NULL ," + _
				"	fas_nome_en TEXT(250) WITH COMPRESSION NULL ," + _
				"	fas_nome_fr TEXT(250) WITH COMPRESSION NULL ," + _
				"	fas_nome_de TEXT(250) WITH COMPRESSION NULL ," + _
				"	fas_nome_es TEXT(250) WITH COMPRESSION NULL ," + _
				"	fas_notti INTEGER NULL " + _
				" ) ; " + _
				" CREATE TABLE Arel_listini_appartamenti ( " + _
				"	rla_id COUNTER CONSTRAINT PK_Arel_listini_appartamenti PRIMARY KEY , " + _
				"	rla_listino_id INT NOT NULL, " + _
				"   rla_appartamento_id INT NOT NULL, " + _
				"	rla_fascia_id INT NOT NULL, " + _
				"	rla_prezzo DOUBLE NULL " + _
				" ) ; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_listini_fasce " + _
				" 	FOREIGN KEY (rla_fascia_id) REFERENCES Atb_listini_fasce(fas_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_listini " + _
				"	FOREIGN KEY (rla_listino_id) REFERENCES Atb_listini(lis_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_appartamenti " + _
				"	FOREIGN KEY (rla_appartamento_id) REFERENCES Atb_appartamenti(ap_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; "
		case DB_SQL
			Aggiornamento__FLAT__2 = _
				"CREATE TABLE Atb_listini (" + _
	  			"	lis_id INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_atb_listini PRIMARY KEY ," + _
	  			"	lis_nome_it nvarchar(250) NULL ," + _
				"	lis_nome_en nvarchar(250) NULL ," + _
				"	lis_nome_fr nvarchar(250) NULL ," + _
				"	lis_nome_de nvarchar(250) NULL ," + _
				"	lis_nome_es nvarchar(250) NULL ," + _
				"	lis_data SMALLDATETIME NULL, " + _
				"	lis_condizioni_it text NULL, " + _
	  			"	lis_condizioni_en text NULL, " + _
				"	lis_condizioni_fr text NULL, " + _
				"	lis_condizioni_es text NULL, " + _
				"	lis_condizioni_de text NULL " + _
			  	" ) ; " + _
				" CREATE TABLE Atb_listini_fasce ( " + _
				"	fas_id INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_atb_listini_fasce PRIMARY KEY , " + _
				"	fas_nome_it nvarchar(250) NULL ," + _
				"	fas_nome_en nvarchar(250) NULL ," + _
				"	fas_nome_fr nvarchar(250) NULL ," + _
				"	fas_nome_de nvarchar(250) NULL ," + _
				"	fas_nome_es nvarchar(250) NULL ," + _
				"	fas_notti INTEGER NULL " + _
				" ) ; " + _
				" CREATE TABLE Arel_listini_appartamenti ( " + _
				"	rla_id INT IDENTITY (1, 1) NOT NULL CONSTRAINT PK_Arel_listini_appartamenti PRIMARY KEY , " + _
				"	rla_listino_id INT NOT NULL, " + _
				"   rla_appartamento_id INT NOT NULL, " + _
				"	rla_fascia_id INT NOT NULL, " + _
				"	rla_prezzo REAL NULL " + _
				" ) ; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_listini_fasce " + _
				" 	FOREIGN KEY (rla_fascia_id) REFERENCES Atb_listini_fasce(fas_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_listini " + _
				"	FOREIGN KEY (rla_listino_id) REFERENCES Atb_listini(lis_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE Arel_listini_appartamenti ADD CONSTRAINT FK__Arel_listini_appartamenti__Atb_appartamenti " + _
				"	FOREIGN KEY (rla_appartamento_id) REFERENCES Atb_appartamenti(ap_id) " + _
				"	ON UPDATE CASCADE ON DELETE CASCADE; "
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FLAT 3
'...........................................................................................
'modifico campo ordine dotazione
'...........................................................................................
function Aggiornamento__FLAT__3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FLAT__3 = "ALTER TABLE atb_prenotazioni ADD "+ _
									 "	pre_chiave TEXT(50) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__FLAT__3 = "ALTER TABLE atb_prenotazioni ADD " + _
									 "	pre_chiave nvarcar(50) NULL;"
	end select
end function
'*******************************************************************************************
%>
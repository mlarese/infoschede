<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../IndexContent/Tools_IndexContent.asp" -->
<% 

'*******************************************************************************************
'AGGIORNAMENTO 1
'...........................................................................................
'aggiunge struttura dati per NextFlat
'...........................................................................................
sql = "CREATE TABLE Atb_Prenotazioni (" &_
	  "		pre_ID COUNTER CONSTRAINT PK_Atb_Prenotazioni PRIMARY KEY, " &_
	  "		pre_NextCom_ID INTEGER NOT NULL, " &_
	  "		pre_ap_ID INTEGER NOT NULL, " &_
	  "		pre_posti INTEGER NULL, " &_
	  "		pre_data_inizio DATETIME NULL, " &_
	  "		pre_data_fine DATETIME NULL " &_
	  ");" &_
	  "CREATE TABLE Atb_Appartamenti (" &_
	  "		ap_ID COUNTER CONSTRAINT PK_Atb_Appartamenti PRIMARY KEY, " &_
	  "		ap_denominazione varchar(250) NULL ," &_
	  "		ap_descrizione_IT TEXT WITH COMPRESSION NULL," &_
	  "		ap_descrizione_EN TEXT WITH COMPRESSION NULL," &_
	  "		ap_descrizione_FR TEXT WITH COMPRESSION NULL," &_
	  "		ap_descrizione_DE TEXT WITH COMPRESSION NULL," &_
	  "		ap_descrizione_ES TEXT WITH COMPRESSION NULL," &_
	  "		ap_NextWeb_ps_mappa INTEGER NULL, " &_
	  "		ap_NextWeb_ps_gallery INTEGER NULL, " &_
	  "		ap_foto varchar(250) NULL ," &_
	  "		ap_minimo_notti INTEGER NULL, " &_
	  "		ap_posti_letto varchar(50) NULL ," &_
	  "		ap_home BIT NULL ," &_
	  "		ap_notti_1 INTEGER NULL, " &_
	  "		ap_sconto_1 INTEGER NULL, " &_
	  "		ap_notti_2 INTEGER NULL, " &_
	  "		ap_sconto_2 INTEGER NULL, " &_
	  "		ap_notti_3 INTEGER NULL, " &_
	  "		ap_sconto_3 INTEGER NULL, " &_
	  "		ap_notti_4 INTEGER NULL, " &_
	  "		ap_sconto_4 INTEGER NULL, " &_
	  "		ap_notti_5 INTEGER NULL, " &_
	  "		ap_sconto_5 INTEGER NULL " &_
	  ");" &_
	  "CREATE TABLE Atb_ap_disponibilita (" &_
	  "		dispo_ID COUNTER CONSTRAINT PK_Atb_ap_disponibilita PRIMARY KEY, " &_
	  "		dispo_ap_ID INTEGER NOT NULL, " &_
	  "		dispo_data DATETIME NULL, " &_
	  "		dispo_disponibile BIT NULL ," &_
	  "		dispo_prezzo CURRENCY NOT NULL " &_
	  ");" &_
	  "CREATE TABLE Atb_ap_dotazioni (" &_
	  "		dot_ID COUNTER CONSTRAINT PK_Atb_ap_dotazioni PRIMARY KEY, " &_
	  "		dot_ap_ID INTEGER NOT NULL, " &_
	  "		dot_ordine char(1) NULL ," &_
	  "		dot_valore_IT TEXT(250) WITH COMPRESSION NULL," &_
	  "		dot_valore_EN TEXT(250) WITH COMPRESSION NULL," &_
	  "		dot_valore_FR TEXT(250) WITH COMPRESSION NULL," &_
	  "		dot_valore_DE TEXT(250) WITH COMPRESSION NULL," &_
	  "		dot_valore_ES TEXT(250) WITH COMPRESSION NULL" &_
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
CALL DB.Execute(sql, 1)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'aggiunge indice su indirizzario
'...........................................................................................
sql = 	  "CREATE INDEX IDX_tb_indirizzario_Owner_ID ON tb_indirizzario (Owner_ID) WITH IGNORE NULL; " &_
		  "CREATE INDEX IDX_tb_indirizzario_Owner_Site ON tb_indirizzario (Owner_Site) WITH IGNORE NULL; "
CALL DB.Execute(sql, 2)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
'...........................................................................................
sql = "CREATE TABLE tb_emailConfig ( " &_
		"config_id COUNTER CONSTRAINT PK_config_id PRIMARY KEY, " &_
		"config_host varchar(250) NOT NULL, " &_
		"config_port integer, " &_
		"config_user varchar(50) NOT NULL, " &_
		"config_pass varchar(50) NOT NULL, " &_
		"config_protocol varchar(5) NOT NULL, " &_
		"config_email varchar(250) NOT NULL, " &_
		"config_deleteMessage BIT NULL, " &_
		"config_delayDelMessage INTEGER NULL, " &_
		"config_id_empl INTEGER NOT NULL " &_
		"); " &_
		"ALTER TABLE tb_emailConfig ADD CONSTRAINT FK_tb_emailConfig__tb_dipendenti " &_
		"FOREIGN KEY (config_id_empl) REFERENCES tb_dipendenti (dip_id) " &_
		"ON UPDATE CASCADE ON DELETE CASCADE; " &_
		"CREATE INDEX IDX_tb_emailConfig_config_id_empl ON tb_emailConfig (config_id_empl)"

CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4
'...........................................................................................
'aggiunge campo per percentuale caparra su tabella appartamenti
'...........................................................................................
sql = "ALTER TABLE Atb_Appartamenti ADD COLUMN ap_caparra integer NULL "
CALL DB.Execute(sql, 4)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 5
'...........................................................................................
'toglie campi inutilizzati su tabella prenotazioni e aggiunge campi: data di arrivo, numero di notti, prezzo totale
'...........................................................................................
sql = "ALTER TABLE Atb_prenotazioni DROP COLUMN pre_posti, pre_data_inizio, pre_data_fine;" &_
	  "ALTER TABLE Atb_prenotazioni ADD COLUMN " &_
	  " pre_dataArrivo DATETIME NULL, " &_
	  " pre_NumeroNotti INTEGER NULL, " &_
	  " pre_prezzo CURRENCY NULL" 
CALL DB.Execute(sql, 5)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 6
'...........................................................................................
'toglie campo inutilizzato per percentuale caparra su tabella appartamenti
'...........................................................................................
sql = "ALTER TABLE Atb_Appartamenti DROP COLUMN ap_caparra"
CALL DB.Execute(sql, 6)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 7
'...........................................................................................
'aggiunge campo per conferma della prenotazione su tabella prenotazioni
'...........................................................................................
sql = "ALTER TABLE Atb_Prenotazioni ADD COLUMN pre_confermata BIT "
CALL DB.Execute(sql, 7)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 8
'...........................................................................................
'aggiunge campi descrizione ai link utili
'...........................................................................................
sql = "ALTER TABLE tb_Links ADD COLUMN " &_
	  "link_descr_IT TEXT WITH COMPRESSION NULL, " &_
	  "link_descr_EN TEXT WITH COMPRESSION NULL, " &_
	  "link_descr_FR TEXT WITH COMPRESSION NULL, " &_
	  "link_descr_DE TEXT WITH COMPRESSION NULL, " &_
	  "link_descr_ES TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 8)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 9
'...........................................................................................
'aggiunge campo per numero di persone a tabella prenotazione
'...........................................................................................
sql = "ALTER TABLE Atb_Prenotazioni ADD COLUMN pre_posti INTEGER "
CALL DB.Execute(sql, 9)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 10
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' e_mail_in BIT =0 posta inviata, email_in = 1 posta ricevuta
'...........................................................................................
sql = "ALTER TABLE tb_email ADD COLUMN email_in BIT; "
CALL DB.Execute(sql, 10)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 11
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' aggiunge MessageID
'...........................................................................................
sql = "ALTER TABLE tb_email ADD COLUMN email_MessageID varchar(100); "
CALL DB.Execute(sql, 11)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 12
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' aggiunge rel_dip_email e relative relazioni
'...........................................................................................
sql = "CREATE TABLE rel_dip_email ( " &_
		"rel_id COUNTER CONSTRAINT PK_rel_id PRIMARY KEY, " &_
		"rel_emailSender varchar(250) NOT NULL, " &_
		"rel_emailSenderID integer, " &_
		"rel_emailID integer, " &_
		"rel_dipID integer, " &_
		"rel_Read BIT NULL, " &_
		"rel_Reply BIT NULL " &_
		"); " &_
		"ALTER TABLE rel_dip_email ADD CONSTRAINT FK_rel_dip_email__tb_dipendenti " &_
		"FOREIGN KEY (rel_dipID) REFERENCES tb_dipendenti (dip_id) " &_
		"ON UPDATE CASCADE ON DELETE CASCADE; " &_
		"ALTER TABLE rel_dip_email ADD CONSTRAINT FK_rel_dip_email__tb_email " &_
		"FOREIGN KEY (rel_emailID) REFERENCES tb_email (email_id) " &_
		"ON UPDATE CASCADE ON DELETE CASCADE; " &_
		"CREATE INDEX IDX_rel_dip_email_rel_emailSenderID ON rel_dip_email (rel_emailSenderID);" &_
		"CREATE INDEX IDX_rel_dip_email_rel_dipID ON rel_dip_email (rel_dipID);" &_
		"CREATE INDEX IDX_rel_dip_email_rel_emailID ON rel_dip_email (rel_emailID);"
CALL DB.Execute(sql, 12)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 13
'...........................................................................................
'toglie campo sito_corrente da tabella siti del next-passport
'...........................................................................................
sql = "ALTER TABLE tb_siti DROP COLUMN sito_corrente; "
CALL DB.Execute(sql, 13)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 14
'...........................................................................................
'aggiunge campo su tabella siti che indica se l'applicazione e' un'area riservata esterna
'...........................................................................................
sql = "ALTER TABLE tb_siti ADD COLUMN sito_amministrazione bit ; " &_
	  "UPDATE tb_siti SET sito_amministrazione=1"
CALL DB.Execute(sql, 14)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 15
'...........................................................................................
'modifica permessi di accesso all'area di amministrazione degli utenti 
'...........................................................................................
sql = "UPDATE tb_siti SET sito_p1='PASS_ADMINISTRATOR', sito_p2='PASS_WEBMASTERS_ADMIN', sito_p3='PASS_USERS_ADMIN' " &_
	  " WHERE sito_ID=1"
CALL DB.Execute(sql, 15)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 16
'...........................................................................................
'aggiunge campo su tabella dipendenti per la scadenza dell'account
'...........................................................................................
sql = "ALTER TABLE tb_dipendenti ADD COLUMN dip_scadenza_account DATETIME NULL"
CALL DB.Execute(sql, 16)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 17
'...........................................................................................
'aggiunge campi su tabella tb_indirizzario per la gestione del blocco di un record da parte 
'	di applicazioni esterne. 
'LockedByApplication	-->	contatore che indica quante applicazioni bloccano il contatto
'ApplicationsLocker		-->	elenco id applicazioni separate da " " che bloccano il contato
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD COLUMN " &_
	  " LockedByApplication INTEGER NULL, " &_ 
	  " ApplicationsLocker TEXT(50) NULL; " &_
	  " UPDATE tb_indirizzario SET ApplicationsLocker = (' ' & Owner_Site & ', '), LockedByApplication=1 WHERE Owner_Site>0; " &_
	  " DROP INDEX IDEsterno ON tb_indirizzario; "
CALL DB.Execute(sql, 17)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 18
'...........................................................................................
'toglie relazione tra tabella indirizzi e siti (utilizzata precedentemente per indicare 
'la proprieta' del record.
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario DROP CONSTRAINT tb_sititb_Indirizzario; " &_
	  " DROP INDEX IDX_tb_indirizzario_Owner_ID ON tb_indirizzario; " &_
	  " DROP INDEX IDX_tb_indirizzario_Owner_Site ON tb_indirizzario; " &_
	  "ALTER TABLE tb_indirizzario DROP COLUMN Owner_ID;" &_
      "ALTER TABLE tb_indirizzario DROP COLUMN Owner_Site;"
CALL DB.Execute(sql, 18)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 19
'...........................................................................................
'aggiunge tabelle per la gestione degli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE tb_Utenti (" & _
	  " 	ut_ID COUNTER CONSTRAINT PK_tb_Utenti PRIMARY KEY, " & _
	  " 	ut_NextCom_ID INTEGER NOT NULL, " & _
	  " 	ut_login TEXT(50) WITH COMPRESSION NULL, " & _
	  " 	ut_password TEXT(50) WITH COMPRESSION NULL, " & _
	  " 	ut_Abilitato bit, " & _
	  " 	ut_ScadenzaAccesso DATETIME NULL ); " &_
	  " ALTER TABLE tb_Utenti ADD CONSTRAINT FK_tb_Utenti__tb_Indirizzario " &_
   	  " 	FOREIGN KEY (ut_NextCom_ID) REFERENCES Tb_Indirizzario (IDElencoIndirizzi) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  " CREATE INDEX IDX_tb_Utenti__ut_NextCom_ID ON tb_Utenti (ut_NextCom_ID);"
CALL DB.Execute(sql, 19)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 20
'...........................................................................................
'aggiunge tabelle per la gestione dei permessi degli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE rel_utenti_sito (" & _
	  "		rel_id COUNTER CONSTRAINT PK_rel_utenti_sito PRIMARY KEY, " & _
	  "		rel_ut_id INTEGER NOT NULL, " & _
	  "		rel_sito_id INTEGER NOT NULL, " & _
	  "		rel_permesso INTEGER NOT NULL ); " &_
	  " ALTER TABLE rel_utenti_sito ADD CONSTRAINT FK_rel_utenti_sito__tb_utenti " &_
   	  " 	FOREIGN KEY (rel_ut_id) REFERENCES tb_Utenti (Ut_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_	
	  " ALTER TABLE rel_utenti_sito ADD CONSTRAINT FK_rel_utenti_sito__tb_siti " &_
   	  " 	FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (sito_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_	
	  " CREATE INDEX IDX_rel_utenti_sito__rel_ut_id ON rel_utenti_sito (rel_ut_id); " &_
	  " CREATE INDEX IDX_rel_utenti_sito__rel_sito_id ON rel_utenti_sito (rel_sito_id); "
CALL DB.Execute(sql, 20)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 21
'...........................................................................................
'aggiunge tabella per log degli accessi per gli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE log_utenti (" & _
	  "		log_id COUNTER CONSTRAINT PK_log_utenti PRIMARY KEY, " & _
	  "		log_ut_id INTEGER NOT NULL, " &_
	  "		log_sito_id INTEGER NOT NULL, " &_
	  "		log_data DATETIME NULL, " &_
	  "		log_username TEXT(50) WITH COMPRESSION NULL ); " &_
	  " ALTER TABLE log_utenti ADD CONSTRAINT FK_log_utenti__tb_utenti " &_
   	  " 	FOREIGN KEY (log_ut_id) REFERENCES tb_Utenti (Ut_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  " ALTER TABLE log_utenti ADD CONSTRAINT FK_log_utenti__tb_siti " &_
   	  " 	FOREIGN KEY (log_sito_id) REFERENCES tb_siti (sito_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " 
CALL DB.Execute(sql, 21)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 22
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' aggiunge MessageID
'...........................................................................................
sql = "ALTER TABLE tb_email ADD COLUMN email_UIDL varchar(250); "
CALL DB.Execute(sql, 22)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 23
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' aggiunge MessageID
'...........................................................................................
sql = "ALTER TABLE tb_email DROP COLUMN email_UIDL; " &_
		"ALTER TABLE tb_email ADD COLUMN email_UIDL INTEGER NOT NULL; "
CALL DB.Execute(sql, 23)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 24
'...........................................................................................
'aggiunge tabelle per la gestione della posta in ingresso
' aggiunge MessageID
'...........................................................................................
sql = "ALTER TABLE tb_email " &_
		"ADD COLUMN email_Account INTEGER NOT NULL, " &_
		" 	email_To TEXT(250) WITH COMPRESSION NULL, " &_
		"	email_CC TEXT(250) WITH COMPRESSION NULL, " &_
		"	email_mime TEXT(50) WITH COMPRESSION NULL, " &_
		"	email_From TEXT(250) WITH COMPRESSION NULL" 
CALL DB.Execute(sql, 24)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 25
'...........................................................................................
'aggiunge campo alle applicazioni per indicare qual'e' la rubrica collegata
'aggiunge campo a rubriche per indicare le rubriche bloccate dal sistema e non utilizzabili 
'nel normale utilizzo dell'indirizzario
'...........................................................................................
sql = "ALTER TABLE tb_siti ADD COLUMN sito_rubrica_area_riservata INTEGER NULL; " &_
	  "ALTER TABLE tb_rubriche ADD COLUMN rubrica_esterna bit "
CALL DB.Execute(sql, 25)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 26
'...........................................................................................
'aggiunge tabella per gestione circolari
'...........................................................................................
sql = "CREATE TABLE tb_Circolari( " & _
	  "	CI_id COUNTER CONSTRAINT PK_tb_circolari PRIMARY KEY, " & _
	  " CI_Numero TEXT(50) WITH COMPRESSION NULL, " & _
	  " CI_Titolo TEXT(250) WITH COMPRESSION NULL, " & _
	  " CI_Estratto TEXT WITH COMPRESSION NULL, " & _
	  " CI_Pubblicazione DATETIME NULL, " & _
	  " CI_Scadenza DATETIME NULL, " & _
	  " CI_File TEXT WITH COMPRESSION NULL, " & _
	  " CI_Visibile BIT, " & _
	  " CI_Protetto BIT )"
CALL DB.Execute(sql, 26)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 27
'...........................................................................................
'aggiunge tabella log dei download dei documenti
'...........................................................................................
sql = " CREATE TABLE log_circolari (" & _
	  "		log_id COUNTER CONSTRAINT PK_log_circolari PRIMARY KEY, " & _
	  "		log_ut_id INTEGER NOT NULL, " &_
	  "		log_dip_id INTEGER NOT NULL, " &_
	  "		log_ci_id INTEGER NOT NULL, " &_
	  "		log_data DATETIME NULL ); " &_
	  " ALTER TABLE log_circolari ADD CONSTRAINT FK_log_circolari__tb_Circolari " &_
   	  " 	FOREIGN KEY (log_ci_id) REFERENCES tb_circolari (CI_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " 
CALL DB.Execute(sql, 27)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 28
'...........................................................................................
'aggiunge tabella tb_admin e ne copia il contentuto dalla tabella dipendenti
'...........................................................................................
sql = " CREATE TABLE tb_admin (" &_
	  " 	id_admin COUNTER CONSTRAINT PK_log_circolari PRIMARY KEY, " & _
	  "		admin_nome TEXT(50) WITH COMPRESSION NULL, " & _
	  "		admin_cognome TEXT(50) WITH COMPRESSION NULL, " & _
	  "		admin_email TEXT(50) WITH COMPRESSION NULL, " & _
	  "		admin_note TEXT WITH COMPRESSION NULL, " & _
	  "		admin_login TEXT(50) WITH COMPRESSION NULL, " & _
	  "		admin_password TEXT(50) WITH COMPRESSION NULL, " & _
	  "		admin_scadenza DATETIME NULL); " &_
	  "	INSERT INTO tb_admin(id_admin, admin_nome, admin_cognome, admin_email, admin_note, admin_login, admin_password, admin_scadenza ) " &_
	  " SELECT dip_id, dip_nome, dip_cognome, dip_email, dip_note, dip_login, dip_password, dip_scadenza_account FROM tb_dipendenti"
CALL DB.Execute(sql, 28)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 29
'...........................................................................................
'cambia relazione tra log di accesso dei dipendenti dalla tabella dei dipendenti a tb_admin
'...........................................................................................
sql = " ALTER TABLE log_admin ADD COLUMN log_admin_ID INTEGER ; " &_
	  " UPDATE log_admin SET log_admin_id=log_dip_id; " &_
	  " ALTER TABLE log_admin DROP CONSTRAINT tb_dipendentilog_admin; " &_ 
	  "	ALTER TABLE log_admin DROP COLUMN log_dip_id; " &_
	  " ALTER TABLE log_admin ADD CONSTRAINT FK_log_admin__tb_admin " &_
	  " FOREIGN KEY (log_admin_id) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 29)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 30
'...........................................................................................
'cambia relazione tra permessi del sito e tabella amministratori
'...........................................................................................
sql = " ALTER TABLE rel_admin_sito ADD COLUMN rel_admin_id INTEGER ; " &_
	  " UPDATE rel_admin_sito SET rel_admin_id=rel_dip_id; " &_
	  " ALTER TABLE rel_admin_sito DROP CONSTRAINT tb_dipendentirel_admin_sito; " &_ 
	  "	ALTER TABLE rel_admin_sito DROP COLUMN rel_dip_id; " &_
	  " ALTER TABLE rel_admin_sito ADD CONSTRAINT FK_rel_admin_sito__tb_admin " &_
	  " FOREIGN KEY (rel_admin_id) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 30)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 31
'...........................................................................................
'cambia relazione tra gruppi di lavoro e dipendenti e la collega con tb_admin
'...........................................................................................
sql = " ALTER TABLE tb_rel_dipgruppi DROP CONSTRAINT [{B026AEDF-A411-4821-83B2-0292BD54EBB7}]; " &_ 
	  " ALTER TABLE tb_rel_dipgruppi ADD CONSTRAINT FK_tb_rel_dipgruppi__tb_admin " &_
	  " FOREIGN KEY (id_impiegato) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 31)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 32
'...........................................................................................
'cambia relazione tra email spedite e tb_dipendenti con tb_admin
'...........................................................................................
sql = " ALTER TABLE rel_dip_email DROP CONSTRAINT FK_rel_dip_email__tb_dipendenti ; " &_ 
	  " ALTER TABLE rel_dip_email ADD CONSTRAINT FK_rel_dip_email__tb_admin " &_
	  " FOREIGN KEY (rel_dipID) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 32)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 33
'...........................................................................................
'cambia relazione tra parametri di configurazione email e tb_dipendenti con tb_admin
'...........................................................................................
sql = " ALTER TABLE tb_EmailConfig DROP CONSTRAINT FK_tb_EmailConfig__tb_dipendenti ; " &_ 
	  " ALTER TABLE tb_EmailConfig ADD CONSTRAINT FK_tb_EmailConfig__tb_admin " &_
	  " FOREIGN KEY (config_id_empl) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 33)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 34
'...........................................................................................
'cancella definitivamente tabella dipendenti
'...........................................................................................
sql = " DROP TABLE tb_dipendenti"
CALL DB.Execute(sql, 34)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 35
'...........................................................................................
'cambia nomi dei campi alla tabella rel_admin_sito e della tabella tb_siti
'...........................................................................................
sql = "SELECT * INTO TMP_rel_admin_sito FROM rel_admin_sito; " &_
	  "ALTER TABLE rel_admin_sito DROP CONSTRAINT FK_rel_admin_sito__tb_admin; " &_
	  "ALTER TABLE rel_admin_sito DROP CONSTRAINT tb_sitirel_admin_sito; " &_
	  "DROP TABLE rel_admin_sito; " &_
	  "CREATE TABLE rel_admin_sito( " &_
	  "		id_p COUNTER CONSTRAINT PK_rel_admin_sito PRIMARY KEY, " &_
	  "		admin_id INTEGER, " &_
	  "		rel_as_permesso INTEGER, " &_
	  "		sito_id INTEGER); " &_
	  "INSERT INTO rel_admin_sito (id_p, admin_id, rel_as_permesso, sito_id) " &_
	  "SELECT rel_id, rel_admin_id, rel_permesso, rel_sito_id FROM TMP_rel_admin_sito; " &_
	  " ALTER TABLE rel_admin_sito ADD CONSTRAINT FK_rel_admin_sito__tb_admin " &_
	  " FOREIGN KEY (admin_id) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  "DROP TABLE TMP_rel_admin_sito; " &_
	  "SELECT * INTO TMP_tb_siti FROM tb_siti; " &_
	  "ALTER TABLE rel_utenti_sito DROP CONSTRAINT FK_rel_utenti_sito__tb_siti; " &_
	  "ALTER TABLE log_utenti DROP CONSTRAINT FK_log_utenti__tb_siti; " &_
	  "ALTER TABLE log_admin DROP CONSTRAINT tb_sitilog_admin; " &_
	  "DROP TABLE tb_siti; " &_
	  "CREATE TABLE tb_siti( " &_
	  "		id_sito INTEGER CONSTRAINT PK_tb_siti PRIMARY KEY, " &_
	  "		sito_nome TEXT(250) WITH COMPRESSION NULL, " &_
	  "		sito_dir TEXT(150) WITH COMPRESSION NULL, " &_
	  "		sito_p1 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p2 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p3 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p4 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p5 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p6 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p7 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p8 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_p9 TEXT(50) WITH COMPRESSION NULL, " &_
	  "		sito_amministrazione bit, " &_
	  "		sito_rubrica_area_riservata INTEGER NULL); " &_
	  "INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_p1, sito_p2, sito_p3, sito_p4, sito_p5, sito_p6, sito_p7, sito_p8, sito_p9, sito_amministrazione, sito_rubrica_area_riservata) " &_
	  "SELECT sito_id, sito_nome, sito_dir, sito_p1, sito_p2, sito_p3, sito_p4, sito_p5, sito_p6, sito_p7, sito_p8, sito_p9, sito_amministrazione, sito_rubrica_area_riservata FROM TMP_tb_siti; " &_
	  "DROP TABLE TMP_tb_siti; " &_
	  " ALTER TABLE rel_utenti_sito ADD CONSTRAINT FK_rel_utenti_sito__tb_siti " &_
   	  " 	FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_	
	  " ALTER TABLE log_utenti ADD CONSTRAINT FK_log_utenti__tb_siti " &_
   	  " 	FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  " ALTER TABLE rel_admin_sito ADD CONSTRAINT FK_rel_admin_sito__tb_siti " &_
   	  " 	FOREIGN KEY (sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_	
	  " ALTER TABLE log_admin ADD CONSTRAINT FK_log_admin__tb_siti " &_
   	  " 	FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 35)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 36
'...........................................................................................
'modifica permessi di accesso all'area di amministrazione degli utenti 
'...........................................................................................
sql = "UPDATE tb_siti SET sito_p1='PASS_ADMIN', sito_p2='PASS_AMMINISTRATORI', sito_p3='PASS_UTENTI' " &_
	  " WHERE ID_sito=1"
CALL DB.Execute(sql, 36)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 37
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom
'...........................................................................................
sql = " ALTER TABLE tb_Indirizzario ADD COLUMN" & _
	  " SyncroKey TEXT(50) WITH COMPRESSION NULL, " & _
	  " SyncroTable TEXT(50) WITH COMPRESSION NULL, " &_
	  " SyncroApplication INTEGER NULL"
CALL DB.Execute(sql, 37)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 38
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom
'...........................................................................................
sql = " ALTER TABLE tb_ValoriNumeri ADD COLUMN SyncroField TEXT(50) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 38)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 39
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_rubriche ADD COLUMN" & _
	  " SyncroTable TEXT(50) WITH COMPRESSION NULL, " &_
	  " SyncroFilter INTEGER NULL"
CALL DB.Execute(sql, 39)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 40
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_indirizzario ADD COLUMN LocalitaElencoIndirizzi TEXT(100) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 40)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 41
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_rubriche ADD " & _
	  " SyncroFilterTable TEXT(50) WITH COMPRESSION NULL, " &_
	  " SyncroFilterKey INTEGER NULL; " &_
	  " ALTER TABLE tb_rubriche DROP COLUMN SyncroFilter; "
CALL DB.Execute(sql, 41)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 42
'...........................................................................................
'crea indici su struttura dati dell'indirizzario
'...........................................................................................
sql = " CREATE INDEX IX_rel_rub_ind ON rel_rub_ind(id_indirizzo); " &_
	  " CREATE INDEX IX_tb_ValoriNumeri ON tb_ValoriNumeri(id_Indirizzario)" 
CALL DB.Execute(sql, 42)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 43
'...........................................................................................
'crea tabelle del nextBanner
'...........................................................................................
sql = "CREATE TABLE tb_banner ("+ _
	  "		ban_id COUNTER CONSTRAINT PK_tb_banner PRIMARY KEY, "+ _
	  "		ban_nome TEXT(50) WITH COMPRESSION NULL, "+ _
	  "		ban_image TEXT(50) WITH COMPRESSION NULL, "+ _
	  "		ban_link TEXT(250) WITH COMPRESSION NULL, "+ _
	  "		ban_alt TEXT WITH COMPRESSION NULL, "+ _
	  "		ban_tipo INTEGER NOT NULL, "+ _
	  "		ban_az INTEGER NOT NULL "+ _
	  ");"+ _
	  "CREATE TABLE tb_tipiBanner ("+ _
	  "		tipoB_id COUNTER CONSTRAINT PK_tb_tipiBanner PRIMARY KEY, "+ _
	  "		tipoB_nome TEXT(50) WITH COMPRESSION NULL "+ _
	  ");"+ _
	  "CREATE TABLE rel_banner_pagine ("+ _
	  "		rbp_id COUNTER CONSTRAINT PK_rel_banner_pagine PRIMARY KEY, "+ _
	  "		rbp_impress_iniz INTEGER NULL, "+ _
	  "		rbp_impress INTEGER NULL, "+ _
	  "		rbp_data_iniz DATETIME NULL, "+ _
	  "		rbp_data_fine DATETIME NULL, "+ _
	  "		rbp_click_iniz INTEGER NULL, "+ _
	  "		rbp_click INTEGER NULL, "+ _
	  "		rbp_pag INTEGER NOT NULL, "+ _
	  "		rbp_banner INTEGER NOT NULL "+ _
	  ");"+ _
	  "CREATE TABLE tb_pagine ("+ _
	  "		pag_id COUNTER(10,1) CONSTRAINT PK_tb_pagine PRIMARY KEY, "+ _
	  "		pag_url TEXT(50) WITH COMPRESSION NULL, "+ _
	  "		pag_cat TEXT(50) WITH COMPRESSION NULL, "+ _
	  "		pag_sito INTEGER NOT NULL " + _
	  ");"+ _
	  "CREATE TABLE tb_applicativi ("+ _
	  "		sito_id COUNTER(2,1) CONSTRAINT PK_tb_applicativi PRIMARY KEY, "+ _
	  "		sito_nome TEXT(250) WITH COMPRESSION NULL, "+ _
	  "		sito_url TEXT(250) WITH COMPRESSION NULL "+ _
	  ");"+ _
	  "ALTER TABLE tb_banner ADD CONSTRAINT FK_tb_banner__tb_indirizzario " & _
   	  "		FOREIGN KEY (ban_az) REFERENCES tb_indirizzario (IDElencoIndirizzi) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  "ALTER TABLE tb_banner ADD CONSTRAINT FK_tb_banner__tb_tipiBanner " & _
   	  "		FOREIGN KEY (ban_tipo) REFERENCES tb_tipiBanner (tipoB_id) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  "ALTER TABLE rel_banner_pagine ADD CONSTRAINT FK_rel_banner_pagine__tb_banner " & _
   	  "		FOREIGN KEY (rbp_banner) REFERENCES tb_banner (ban_id) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  "ALTER TABLE rel_banner_pagine ADD CONSTRAINT FK_rel_banner_pagine__tb_pagine " & _
   	  "		FOREIGN KEY (rbp_pag) REFERENCES tb_pagine (pag_id) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  "ALTER TABLE tb_pagine ADD CONSTRAINT FK_tb_pagine__tb_applicativi " & _
   	  "		FOREIGN KEY (pag_sito) REFERENCES tb_applicativi (sito_id) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 43)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 44
'...........................................................................................
'crea tabelle e relazione del nextCRM
'...........................................................................................
sql = "CREATE TABLE tb_pratiche (" + _
	  "pra_id COUNTER CONSTRAINT PK_tb_pratiche PRIMARY KEY ," + _
	  "pra_codice TEXT(50) WITH COMPRESSION NULL ," + _
	  "pra_nome TEXT(255) WITH COMPRESSION NULL ," + _
	  "pra_dataI DATETIME NULL ," + _
	  "pra_dataUM DATETIME NULL ," + _
	  "pra_dataA DATETIME NULL ," + _
	  "pra_archiviata BIT NULL DEFAULT 0 ," + _
	  "pra_note TEXT WITH COMPRESSION ," + _
	  "pra_pubblica BIT NULL ," + _
	  "pra_cliente_id INTEGER NULL ," + _
	  "pra_creatore_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE tb_documenti (" + _
	  "doc_id COUNTER CONSTRAINT PK_tb_documenti PRIMARY KEY ," + _
	  "doc_nome TEXT(255) WITH COMPRESSION NULL ," + _
	  "doc_path TEXT(255) WITH COMPRESSION NULL ," + _
	  "doc_dataC DATETIME NULL ," + _
	  "doc_pubblica BIT NULL DEFAULT 0 ," + _
	  "doc_eredita BIT NULL DEFAULT 1 ," + _
	  "doc_note TEXT WITH COMPRESSION ," + _
	  "doc_tipologia_id INTEGER NULL ," + _
	  "doc_pratica_id INTEGER NULL ," + _
	  "doc_creatore_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE tb_attivita (" + _
	  "att_id COUNTER CONSTRAINT PK_tb_attivita PRIMARY KEY ," + _
	  "att_oggetto TEXT(255) WITH COMPRESSION NULL ," + _
	  "att_testo TEXT WITH COMPRESSION NULL ," + _
	  "att_note TEXT WITH COMPRESSION NULL ," + _
	  "att_dataCrea DATETIME NULL ," + _
	  "att_dataChiusa DATETIME NULL ," + _
	  "att_dataS DATETIME NULL ," + _
	  "att_priorita BIT NULL DEFAULT 0 ," + _
	  "att_conclusa BIT NULL DEFAULT 0 ," + _
	  "att_pubblica BIT NULL ," + _
	  "att_eredita BIT NULL ," + _
	  "att_sistema BIT NULL DEFAULT 0 ," + _
	  "att_domanda_id INTEGER NULL ," + _
	  "att_mittente_id INTEGER NULL ," + _
	  "att_pratica_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE al_attivita_gruppi (" + _
	  "al_id COUNTER CONSTRAINT PK_al_attivita_gruppi PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_gruppo_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_attivita_utenti (" + _
	  "al_id COUNTER CONSTRAINT PK_al_attivita_utenti PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_utente_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_default_gruppi (" + _
	  "al_id COUNTER CONSTRAINT PK_al_default_gruppi PRIMARY KEY ," + _
	  "al_gruppo_id INTEGER NULL ," + _
	  "al_tipo_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_default_utenti (" + _
	  "al_id COUNTER CONSTRAINT PK_al_default_utenti PRIMARY KEY ," + _
	  "al_utente_id INTEGER NULL ," + _
	  "al_tipo_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_documenti_gruppi (" + _
	  "al_id COUNTER CONSTRAINT PK_al_documenti_gruppi PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_gruppo_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_documenti_utenti (" + _
	  "al_id COUNTER CONSTRAINT PK_al_documenti_utenti PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_utente_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_pratiche_gruppi (" + _
	  "al_id COUNTER CONSTRAINT PK_al_pratiche_gruppi PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_gruppo_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE al_pratiche_utenti (" + _
	  "al_id COUNTER CONSTRAINT PK_al_pratiche_utenti PRIMARY KEY ," + _
	  "al_tipo_id INTEGER NULL ," + _
	  "al_utente_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE tb_tipologie (" + _
	  "tipo_id COUNTER CONSTRAINT PK_tb_tipologie PRIMARY KEY ," + _
	  "tipo_nome TEXT(50) WITH COMPRESSION NULL" + _
	  ");" + _
	  "CREATE TABLE tb_descrittori (" + _
	  "descr_id COUNTER CONSTRAINT PK_tb_descrittori PRIMARY KEY ," + _
	  "descr_nome TEXT(50) WITH COMPRESSION NULL ," + _
	  "descr_tipo INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE tb_allegati (" + _
	  "all_id COUNTER CONSTRAINT PK_tb_allegati PRIMARY KEY ," + _
	  "all_attivita_id INTEGER NULL ," + _
	  "all_documento_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE rel_tipologie_descrittori (" + _
	  "rtd_id COUNTER CONSTRAINT PK_rel_tipologie_descrittori PRIMARY KEY ," + _
	  "rtd_tipologia_id INTEGER NULL ," + _
	  "rtd_descrittore_id INTEGER NULL" + _
	  ");" + _
	  "CREATE TABLE rel_documenti_descrittori (" + _
	  "rdd_id COUNTER CONSTRAINT PK_rel_documenti_descrittori PRIMARY KEY ," + _
	  "rdd_valore TEXT(255) WITH COMPRESSION NULL ," + _
	  "rdd_documento_id INTEGER NULL ," + _
	  "rdd_descrittore_id INTEGER NULL" + _
	  ");" + _
	  "ALTER TABLE rel_tipologie_descrittori ADD CONSTRAINT FK_rel_tipologie_descrittori__tb_descrittori " + _
   	  "FOREIGN KEY (rtd_descrittore_id) REFERENCES tb_descrittori (descr_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE rel_tipologie_descrittori ADD CONSTRAINT FK_rel_tipologie_descrittori__tb_tipologie " + _
   	  "FOREIGN KEY (rtd_tipologia_id) REFERENCES tb_tipologie (tipo_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE rel_documenti_descrittori ADD CONSTRAINT FK_rel_documenti_descrittori__tb_documenti " + _
   	  "FOREIGN KEY (rdd_documento_id) REFERENCES tb_documenti (doc_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_documenti ADD CONSTRAINT FK_tb_documenti__tb_tipologie " + _
   	  "FOREIGN KEY (doc_tipologia_id) REFERENCES tb_tipologie (tipo_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_documenti ADD CONSTRAINT FK_tb_documenti__tb_pratiche " + _
   	  "FOREIGN KEY (doc_pratica_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_allegati ADD CONSTRAINT FK_tb_allegati__tb_documenti " + _
   	  "FOREIGN KEY (all_documento_id) REFERENCES tb_documenti (doc_id);" + _
	  "ALTER TABLE tb_allegati ADD CONSTRAINT FK_tb_allegati__tb_attivita " + _
   	  "FOREIGN KEY (all_attivita_id) REFERENCES tb_attivita (att_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_attivita ADD CONSTRAINT FK_tb_attivita__tb_pratiche " + _
   	  "FOREIGN KEY (att_pratica_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_attivita_utenti ADD CONSTRAINT FK_al_attivita_utenti__tb_attivita " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_attivita (att_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_attivita_gruppi ADD CONSTRAINT FK_al_attivita_gruppi__tb_attivita " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_attivita (att_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_pratiche_utenti ADD CONSTRAINT FK_al_pratiche_utenti__tb_pratiche " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_pratiche_gruppi ADD CONSTRAINT FK_al_pratiche_gruppi__tb_pratiche " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_default_utenti ADD CONSTRAINT FK_al_default_utenti__tb_pratiche " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_default_gruppi ADD CONSTRAINT FK_al_default_gruppi__tb_pratiche " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_documenti_utenti ADD CONSTRAINT FK_al_documenti_utenti__tb_documenti " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_documenti (doc_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_documenti_gruppi ADD CONSTRAINT FK_al_documenti_gruppi__tb_documenti " + _
   	  "FOREIGN KEY (al_tipo_id) REFERENCES tb_documenti (doc_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_documenti_utenti ADD CONSTRAINT FK_al_documenti_utenti__tb_admin " + _
   	  "FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_documenti_gruppi ADD CONSTRAINT FK_al_documenti_gruppi__tb_gruppi " + _
   	  "FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_gruppo) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_pratiche_utenti ADD CONSTRAINT FK_al_pratiche_utenti__tb_admin " + _
   	  "FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_pratiche_gruppi ADD CONSTRAINT FK_al_pratiche_gruppi__tb_gruppi " + _
   	  "FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_gruppo) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_default_utenti ADD CONSTRAINT FK_al_default_utenti__tb_admin " + _
   	  "FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_default_gruppi ADD CONSTRAINT FK_al_default_gruppi__tb_gruppi " + _
   	  "FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_gruppo) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_attivita_utenti ADD CONSTRAINT FK_al_attivita_utenti__tb_admin " + _
   	  "FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE al_attivita_gruppi ADD CONSTRAINT FK_al_attivita_gruppi__tb_gruppi " + _
   	  "FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_gruppo) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_pratiche ADD CONSTRAINT FK_tb_pratiche__tb_indirizzario " + _
   	  "FOREIGN KEY (pra_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_pratiche ADD CONSTRAINT FK_tb_pratiche__tb_admin " + _
   	  "FOREIGN KEY (pra_creatore_id) REFERENCES tb_admin (id_admin);" + _
	  "ALTER TABLE tb_documenti ADD CONSTRAINT FK_tb_documenti__tb_admin " + _
   	  "FOREIGN KEY (doc_creatore_id) REFERENCES tb_admin (id_admin);" + _
	  "ALTER TABLE tb_attivita ADD CONSTRAINT FK_tb_attivita__tb_admin " + _
   	  "FOREIGN KEY (att_mittente_id) REFERENCES tb_admin (id_admin);"
CALL DB.Execute(sql, 44)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 45
'...........................................................................................
'aggiunge campo contatore pratiche e suffisso su tabella indirizzario
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD COLUMN " &_
	  " PraticaCount INTEGER NULL DEFAULT 0, " &_
	  " PraticaPrefisso TEXT(5) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 45)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 46
'...........................................................................................
'aggiunge campo contatore pratiche e suffisso su tabella indirizzario
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ALTER COLUMN " &_
	  " ApplicationsLocker TEXT(255) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 46)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 47
'...........................................................................................
'aggiunge campo contatore pratiche e suffisso su tabella indirizzario
'...........................................................................................
sql = "ALTER TABLE Atb_Appartamenti ADD COLUMN " &_
	  " ap_ordine INTEGER NULL"
CALL DB.Execute(sql, 47)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 48
'...........................................................................................
'aggiunge il permesso COM_POWER per il CRM
'...........................................................................................
sql = "UPDATE tb_siti SET sito_p3='COM_POWER' WHERE id_sito=3"
CALL DB.Execute(sql, 48)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 49
'...........................................................................................
'crea tabelle e relazione del nextClub
'...........................................................................................
sql = "CREATE TABLE ctb_soci (" + _
	  "soc_id COUNTER CONSTRAINT PK_ctb_soci PRIMARY KEY ," + _
	  "soc_numero TEXT(20) WITH COMPRESSION NULL ," + _
	  "soc_dataI DATETIME NULL ," + _
	  "soc_tessera INTEGER NULL ," + _
	  "soc_nick TEXT(20) WITH COMPRESSION NULL ," + _
	  "soc_ind_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE ctb_pagamenti (" + _
	  "paga_id COUNTER CONSTRAINT PK_ctb_pagamenti PRIMARY KEY ," + _
	  "paga_importo CURRENCY NULL ," + _
	  "paga_data DATETIME NULL ," + _
	  "paga_operatore TEXT(50) WITH COMPRESSION NULL ," + _
	  "paga_socio_id INTEGER NULL, " + _
	  "paga_tipo_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE ctb_tipiPagamento (" + _
	  "tipP_id COUNTER CONSTRAINT PK_ctb_tipoPagamento PRIMARY KEY ," + _
	  "tipP_nome TEXT(255) WITH COMPRESSION NULL" + _
	  "); " + _
	  "ALTER TABLE ctb_soci ADD CONSTRAINT FK_ctb_soci__tb_indirizzario " + _
   	  "FOREIGN KEY (soc_ind_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE ctb_pagamenti ADD CONSTRAINT FK_ctb_pagamenti__ctb_soci " + _
   	  "FOREIGN KEY (paga_socio_id) REFERENCES ctb_soci (soc_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE ctb_pagamenti ADD CONSTRAINT FK_ctb_pagamenti__ctb_tipiPagamento " + _
   	  "FOREIGN KEY (paga_tipo_id) REFERENCES ctb_tipiPagamento (tipP_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE tb_indirizzario ADD COLUMN " &_
	  " LuogoNascita TEXT(255) WITH COMPRESSION NULL ," + _
	  " CF TEXT(16) WITH COMPRESSION NULL;"
CALL DB.Execute(sql, 49)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 50
'...........................................................................................
'crea tabelle e relazione del nextBooking
'...........................................................................................
sql = "CREATE TABLE btb_tipiCamera (" + _
	  "tipC_id COUNTER CONSTRAINT PK_btb_tipiCamera PRIMARY KEY ," + _
	  "tipC_nome_ita TEXT(50) WITH COMPRESSION NULL ," + _
	  "tipC_nome_eng TEXT(50) WITH COMPRESSION NULL ," + _
	  "tipC_nome_fra TEXT(50) WITH COMPRESSION NULL ," + _
	  "tipC_nome_spa TEXT(50) WITH COMPRESSION NULL ," + _
	  "tipC_nome_ted TEXT(50) WITH COMPRESSION NULL" + _
	  "); " + _
	  "CREATE TABLE btb_disponibilita (" + _
	  "dis_id COUNTER CONSTRAINT PK_btb_disponibilita PRIMARY KEY ," + _
	  "dis_prezzo CURRENCY NULL ," + _
	  "dis_data DATETIME NULL ," + _
	  "dis_disponibilita INTEGER NULL ," + _
	  "dis_tipo_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE btb_listini (" + _
	  "lis_id COUNTER CONSTRAINT PK_btb_listini PRIMARY KEY ," + _
	  "lis_nome TEXT(100) WITH COMPRESSION NULL ," + _
	  "lis_data DATETIME NULL" + _
	  "); " + _
	  "CREATE TABLE btb_listini_tipiCamera (" + _
	  "rlt_id COUNTER CONSTRAINT PK_btb_listini_tipiCamera PRIMARY KEY ," + _
	  "rlt_prezzo CURRENCY NULL ," + _
	  "rlt_listino_id INTEGER NULL ," + _
	  "rlt_tipo_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE btb_prenotazioni (" + _
	  "pre_id COUNTER CONSTRAINT PK_btb_prenotazioni PRIMARY KEY ," + _
	  "pre_data DATETIME NULL ," + _
	  "pre_data_inizio DATETIME NULL ," + _
	  "pre_data_fine DATETIME NULL ," + _
	  "pre_note TEXT WITH COMPRESSION NULL ," + _
	  "pre_cliente_id INTEGER NULL" + _
	  "); " + _
	  "CREATE TABLE btb_prenotazioni_tipiCamera (" + _
	  "rpt_id COUNTER CONSTRAINT PK_btb_prenotazioni_tipiCamera PRIMARY KEY ," + _
	  "rpt_prezzo CURRENCY NULL ," + _
	  "rpt_numero INTEGER NULL ," + _
	  "rpt_prenotazione_id INTEGER NULL ," + _
	  "rpt_tipo_id INTEGER NULL" + _
	  "); " + _
	  "ALTER TABLE btb_disponibilita ADD CONSTRAINT FK_btb_disponibilita__btb_tipiCamera " + _
   	  "FOREIGN KEY (dis_tipo_id) REFERENCES btb_tipiCamera (tipC_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE btb_listini_tipiCamera ADD CONSTRAINT FK_btb_listini_tipiCamera__btb_tipiCamera " + _
   	  "FOREIGN KEY (rlt_tipo_id) REFERENCES btb_tipiCamera (tipC_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE btb_listini_tipiCamera ADD CONSTRAINT FK_btb_listini_tipiCamera__btb_listini " + _
   	  "FOREIGN KEY (rlt_listino_id) REFERENCES btb_listini (lis_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE btb_prenotazioni_tipiCamera ADD CONSTRAINT FK_btb_prenotazioni_tipiCamera__btb_tipiCamera " + _
   	  "FOREIGN KEY (rpt_tipo_id) REFERENCES btb_tipiCamera (tipC_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;" + _
	  "ALTER TABLE btb_prenotazioni_tipiCamera ADD CONSTRAINT FK_btb_prenotazioni_tipiCamera__btb_prenotazioni " + _
   	  "FOREIGN KEY (rpt_prenotazione_id) REFERENCES btb_prenotazioni (pre_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 50)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 51
'...........................................................................................
'aggiunge tabella per gestione guestbook
'...........................................................................................
sql = " CREATE TABLE tb_guestbook (" & _
	  " 	IdGuest COUNTER CONSTRAINT PK_tb_guestbook PRIMARY KEY, " & _
	  " 	Data DATETIME NULL, " & _
	  " 	Visibile BIT NOT NULL, " & _
	  " 	Id_contatto INTEGER NOT NULL, " & _
	  " 	Messaggio TEXT NULL ); " &_
	  " ALTER TABLE tb_guestbook ADD CONSTRAINT FK_tb_guestbook__tb_Indirizzario " &_
   	  " 	FOREIGN KEY (Id_contatto) REFERENCES Tb_Indirizzario (IDElencoIndirizzi) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  " CREATE INDEX IDX_tb_guestbook__ut_NextCom_ID ON tb_guestbook (Id_contatto);"
CALL DB.Execute(sql, 51)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 52
'...........................................................................................
'aggiunge campi per NextClub
'...........................................................................................
sql = "ALTER TABLE ctb_pagamenti ADD COLUMN " &_
	  " paga_pagato BIT; " & _
	  "ALTER TABLE ctb_soci ADD COLUMN " &_
	  " soc_tessera_inviata BIT; "
CALL DB.Execute(sql, 52)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 53
'...........................................................................................
'aggiunge campi per NextBooking
'...........................................................................................
sql = "ALTER TABLE btb_tipiCamera ADD COLUMN " &_
	  " tipC_ordine INTEGER NULL;"
CALL DB.Execute(sql, 53)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 54
'...........................................................................................
'rinomina campi per NextBooking
'...........................................................................................
sql = "ALTER TABLE btb_tipiCamera ADD COLUMN " &_
	  " tipC_nome_it TEXT(50) WITH COMPRESSION NULL ," & _
	  " tipC_nome_en TEXT(50) WITH COMPRESSION NULL ," & _
	  " tipC_nome_fr TEXT(50) WITH COMPRESSION NULL ," & _
	  " tipC_nome_es TEXT(50) WITH COMPRESSION NULL ," & _
	  " tipC_nome_de TEXT(50) WITH COMPRESSION NULL;" & _
	  "ALTER TABLE btb_listini ADD COLUMN " &_
	  " lis_nome_it TEXT(100) WITH COMPRESSION NULL ," & _
	  " lis_nome_en TEXT(100) WITH COMPRESSION NULL ," & _
	  " lis_nome_fr TEXT(100) WITH COMPRESSION NULL ," & _
	  " lis_nome_es TEXT(100) WITH COMPRESSION NULL ," & _
	  " lis_nome_de TEXT(100) WITH COMPRESSION NULL;" & _
	  "UPDATE btb_tipiCamera SET" & _
	  " tipC_nome_it = tipC_nome_ita, " & _
	  " tipC_nome_en = tipC_nome_eng, " & _
	  " tipC_nome_fr = tipC_nome_fra, " & _
	  " tipC_nome_es = tipC_nome_spa, " & _
	  " tipC_nome_de = tipC_nome_ted;" & _
	  "UPDATE btb_listini SET" & _
	  " lis_nome_it = lis_nome;" & _
	  "ALTER TABLE btb_tipiCamera DROP COLUMN tipC_nome_ita, tipC_nome_eng, tipC_nome_fra, " & _
	  "tipC_nome_spa, tipC_nome_ted;" & _
	  "ALTER TABLE btb_listini DROP COLUMN lis_nome;"
CALL DB.Execute(sql, 54)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 55
'...........................................................................................
'aggiunge campi per NextBooking
'...........................................................................................
sql = "ALTER TABLE btb_prenotazioni ADD COLUMN " &_
	  " pre_nomeCC TEXT(255) WITH COMPRESSION NULL ," & _
	  " pre_numeroCC TEXT(255) WITH COMPRESSION NULL ," & _
	  " pre_dataCC DATETIME NULL ," & _
	  " pre_tipoCC TEXT(50) WITH COMPRESSION NULL ;"
CALL DB.Execute(sql, 55)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 56
'...........................................................................................
'modifica campi per NextBooking
'...........................................................................................
sql = "ALTER TABLE btb_prenotazioni ADD COLUMN " &_
	  " pre_totale INTEGER NULL ;" & _
	  "ALTER TABLE btb_prenotazioni_tipiCamera DROP COLUMN rpt_prezzo;"
CALL DB.Execute(sql, 56)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 57
'...........................................................................................
'modifica campi per next-Flat per permettere sconto su periodo notti (percentuale reale)
'...........................................................................................
sql = "ALTER TABLE Atb_Appartamenti ALTER COLUMN ap_sconto_1 DOUBLE NULL ;" & _
	  "ALTER TABLE Atb_Appartamenti ALTER COLUMN ap_sconto_2 DOUBLE NULL ;" & _
	  "ALTER TABLE Atb_Appartamenti ALTER COLUMN ap_sconto_3 DOUBLE NULL ;" & _
	  "ALTER TABLE Atb_Appartamenti ALTER COLUMN ap_sconto_4 DOUBLE NULL ;" & _
	  "ALTER TABLE Atb_Appartamenti ALTER COLUMN ap_sconto_5 DOUBLE NULL ;"
CALL DB.Execute(sql, 57)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 58
'...........................................................................................
'modifica campi per next-Flat per permettere sconto su periodo notti (percentuale reale)
'...........................................................................................
sql = "ALTER TABLE Atb_Appartamenti ADD COLUMN ap_posti_letto_tmp TEXT(10) NULL, ap_posti_letto_agg integer NULL;" & _
	  "UPDATE Atb_Appartamenti SET ap_posti_letto_tmp=ap_posti_letto;" &_
	  "ALTER TABLE Atb_Appartamenti DROP COLUMN ap_posti_letto;" &_
	  "ALTER TABLE Atb_Appartamenti ADD COLUMN ap_posti_letto INTEGER; " &_
	  "UPDATE Atb_Appartamenti SET ap_posti_letto=4; " &_
	  "UPDATE Atb_Appartamenti SET ap_posti_letto=2 WHERE ap_posti_letto_tmp LIKE '2%';" & _
	  "UPDATE Atb_Appartamenti SET ap_posti_letto_agg=1 WHERE ap_posti_letto_tmp LIKE '%1'; " &_
	  "UPDATE Atb_Appartamenti SET ap_posti_letto_agg=2 WHERE ap_posti_letto_tmp LIKE '%2'; " &_
	   "ALTER TABLE Atb_Appartamenti DROP COLUMN ap_posti_letto_tmp;"
CALL DB.Execute(sql, 58)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 59
'...........................................................................................
'aggiunge il campo relazione per l'indirizzario per gestione cnt interni
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD COLUMN " &_
	  " cntRel INTEGER NULL"
CALL DB.Execute(sql, 59)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 60
'...........................................................................................
'aggiunge il campo default per le pratiche causa gestione documenti
'...........................................................................................
sql = "ALTER TABLE tb_pratiche ADD COLUMN " &_
	  " pra_default BIT NULL"
CALL DB.Execute(sql, 60)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 61
'...........................................................................................
'Toglie campo flag per pratica default e toglie relazioni tra attivita' e pratiche 
' e documenti e pratiche
'...........................................................................................
sql = "ALTER TABLE tb_pratiche DROP COLUMN pra_default; " & vbCrLf & _
	  "ALTER TABLE tb_attivita DROP CONSTRAINT FK_tb_attivita__tb_pratiche; " & vbCrLf & _
	  "ALTER TABLE tb_documenti DROP CONSTRAINT FK_tb_documenti__tb_pratiche " & vbCrLf
CALL DB.Execute(sql, 61)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 62
'...........................................................................................
'Aggiunge tabelle e relazioni per applicativo next-Contract
'...........................................................................................
sql = "CREATE TABLE Ctb_bandi ( " & vbCrLF & _
	  " 	bando_id COUNTER CONSTRAINT PK_Ctb_Bandi PRIMARY KEY, " & vbCrLF & _
	  "		bando_pubblicazione DATETIME NULL, " & vbCrLF & _
	  "		bando_scadenza DATETIME NULL, " & vbCrLF & _
	  "		bando_InizioLavori DATETIME NULL, " & vbCrLF & _
	  "		bando_AperturaOfferte DATETIME NULL, " & vbCrLF & _
	  "		bando_titolo_IT TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_titolo_EN TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_testo_IT TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_testo_EN TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_categoria_id INTEGER NOT NULL, " & vbCrLF & _
	  "		bando_BaseAsta TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Ditte INTEGER NULL, " & vbCrLF & _
	  "		bando_Ditta1 TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Ditta2 TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Ribasso TEXT(50) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Numero TEXT(50) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Codice TEXT(50) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_Commessa TEXT(50) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		bando_pubblicato BIT NULL, " & vbCrLf & _
	  "		bando_assegnato BIT NULL " & vbCrLF & _
	  "		);" & vbCrLF & _
	  "CREATE TABLE Ctb_categorie (" & vbCrLf &_
	  "		categoria_id COUNTER CONSTRAINT PK_Ctb_Categorie PRIMARY KEY , " & vbCrLf & _
	  "		categoria_nome_IT TEXT(250) WITH COMPRESSION NULL, " & vbCrLF & _
	  "		categoria_nome_EN TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		categoria_mail_attiva BIT NULL, " & vbCrLf & _
	  "		categoria_mail_testo_prima TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		categoria_mail_testo_dopo TEXT WITH COMPRESSION NULL " & vbCrLF & _
	  "		);" & vbCrLF & _
	  "CREATE TABLE Ctb_DocPdf (" & vbCrLf & _
	  "		pdf_id COUNTER CONSTRAINT PK_Ctb_DocPdf PRIMARY KEY, " & VbCrLf & _
	  "		pdf_bando_id INTEGER NOT NULL, " & vbCrLf & _
	  "		pdf_titolo_IT TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		pdf_titolo_EN TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		pdf_file_IT TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		pdf_file_EN TEXT(250) WITH COMPRESSION NULL " & vbCrLf & _
	  "		);" & vbCrLf & _
	  "CREATE TABLE Ctb_log_download (" & vbCrLF & _
	  "		log_id COUNTER CONSTRAINT PK_Ctb_log_download PRIMARY KEY, " &_
	  "		log_data DATETIME NULL, " & vbCrLF & _
	  "		log_pdf_id INTEGER NOT NULL, " & vbCrLF & _
	  "		log_ut_id INTEGER NOT NULL " & vbCrLf & _
	  "		);" & vbCrLF & _
 	  "CREATE TABLE Ctb_log_mailing (" & vbCrLf & _
	  "		log_id COUNTER CONSTRAINT PK_Ctb_log_mailing PRIMARY KEY, " &_
	  "		log_data DATETIME NULL, " & vbCrLF & _
	  "		log_email TEXT(200) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		log_bando_id INTEGER NOT NULL, " & vbCrLf & _
	  "		log_mail_id INTEGER NOT NULL " & vbCrLF & _
	  "		);" & vbCrLF & _
	  "CREATE TABLE Ctb_Mail_Bandi(" & vbCrLf & _
	  "		mail_id COUNTER CONSTRAINT PK_Ctb_Mail_Bandi PRIMARY KEY, " & _
	  "		mail_ut_id INTEGER NOT NULL, " & vbCrLF & _
	  "		mail_categoria_id INTEGER NOT NULL, " & vbCrLf & _
	  " 	mail_data DATETIME NULL " & vbCrLf & _
	  "		);" & vbCrLF & _
	  "ALTER TABLE Ctb_Bandi ADD CONSTRAINT FK_Ctb_bandi__Ctb_categorie " & vbCrLF & _
	  "		FOREIGN KEY (bando_Categoria_id) REFERENCES Ctb_categorie(categoria_id) " & vbCrLf & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCrLf & _
	  "ALTER TABLE Ctb_DocPdf ADD CONSTRAINT FK_Ctb_DocPdf__Ctb_Bandi " & vbCrLf & _
	  "		FOREIGN KEY (pdf_bando_id) REFERENCES Ctb_bandi(bando_id) " & vbCrLF & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCrLF & _
	  "ALTER TABLE Ctb_log_download ADD CONSTRAINT FK_Ctb_log_download__Ctb_DocPdf " & vbCrLF & _
	  "		FOREIGN KEY (log_pdf_id) REFERENCES Ctb_DocPdf(pdf_id) " & vbCrLf & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCRLf & _
	  "ALTER TABLE Ctb_log_download ADD CONSTRAINT FK_Ctb_log_download__tb_utenti " & vbCrLf & _
	  "		FOREIGN KEY (log_ut_id) REFERENCES tb_utenti(ut_id) " & vbCrLf & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCrLf & _
	  "ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__Ctb_bandi " & vbCrLf & _
	  "		FOREIGN KEY (log_bando_id) REFERENCES Ctb_Bandi(bando_id) " & vbCrLF & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCRLf & _
	  "ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__Ctb_Mail_Bandi " & vbCrLf & _
	  "		FOREIGN KEY (log_mail_id) REFERENCES Ctb_mail_Bandi(mail_id) " & vbCrLF & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCrLF & _
	  "ALTER TABLE Ctb_Mail_Bandi ADD CONSTRAINT FK_Ctb_mail_bandi__Ctb_Categorie " & vbCrLf & _
	  "		FOREIGN KEY (mail_categoria_id) REFERENCES Ctb_categorie(categoria_id) " & vbCrLF & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & vbCrLF & _
	  "ALTER TABLE Ctb_Mail_Bandi ADD CONSTRAINT FK_Ctb_mail_bandi__tb_utenti " & vbCrLf & _
	  "		FOREIGN KEY (mail_ut_id) REFERENCES tb_utenti(ut_id) " & vbCrLF & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE " & vbCrLF
CALL DB.Execute(sql, 62)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 63
'...........................................................................................
'aggiunge campo per la gestione della lingua nella mailing list
'...........................................................................................
sql = "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_prima; " & vbcrLF & _
	  "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_dopo; " & vbcRLf & _
	  "ALTER TABLE Ctb_categorie ADD COLUMN " & vbCrLf & _
	  "		categoria_mail_testo_prima_IT TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		categoria_mail_testo_prima_EN TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		categoria_mail_testo_dopo_IT TEXT WITH COMPRESSION NULL, " & vbCrLF & _
	  "		categoria_mail_testo_dopo_EN TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 63)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 64
'...........................................................................................
'aggiunge gestione lingua su nextCom
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD COLUMN lingua TEXT(2) WITH COMPRESSION NULL; " & vbcrLF & _
	  "UPDATE tb_indirizzario SET lingua = 'it' "
CALL DB.Execute(sql, 64)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 65
'...........................................................................................
'aggiunge gestione lingua su nextCom
'...........................................................................................
sql = "CREATE TABLE tb_cnt_lingue (" & vbCrLF & _
	  "		lingua_codice TEXT(2) WITH COMPRESSION NOT NULL CONSTRAINT PK_tb_cnt_lingue PRIMARY KEY, " & vbCrLF & _
	  "		lingua_nome_IT TEXT(20) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		lingua_nome TEXT(20) WITH COMPRESSION NULL " & vbCRLF & _
	  "		); " & vbCRLF & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('it', 'Italiano', 'Italiano'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('en', 'Inglese', 'English'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('fr', 'Francese', 'Français'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('de', 'Tedesco', 'Deutsch'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('es', 'Spagnolo', 'Español'); " & vbCrLf & _
	  "ALTER TABLE tb_indirizzario ADD CONSTRAINT FK_tb_indirizzario__tb_cnt_lingue " & vbCrLf &_
	  "		FOREIGN KEY (lingua) REFERENCES tb_cnt_lingue(lingua_codice) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 65)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 66
'...........................................................................................
'modifica gestione mailing list per categorie
'...........................................................................................
sql = "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_prima_EN; " & _
	  "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_prima_IT; " & _
	  "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_dopo_EN; " & _
	  "ALTER TABLE Ctb_categorie DROP COLUMN categoria_mail_testo_dopo_IT; " & _
	  "ALTER TABLE Ctb_categorie ADD categoria_NEXTWEB_ps_mail int NULL;"
CALL DB.Execute(sql, 66)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 67
'...........................................................................................
'modifica gestione log di spedizione email: viene incorporato nel next-com
'...........................................................................................
sql = " ALTER TABLE Ctb_log_mailing DROP CONSTRAINT FK_Ctb_log_mailing__Ctb_Mail_Bandi; " & _
	  " ALTER TABLE Ctb_log_mailing DROP COLUMN log_mail_id; " & _
	  " ALTER TABLE Ctb_log_mailing DROP COLUMN log_email; " & _
	  "	ALTER TABLE Ctb_log_mailing ADD log_categoria_id INT NOT NULL, " &_
	  "		log_NEXTCOM_email_id INT NOT NULL ;" & _
	  " ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__ctb_categoria " & _
	  "		FOREIGN KEY (log_categoria_id) REFERENCES ctb_categorie(categoria_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ;" & _
	  " ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__tb_email " & _
	  "		FOREIGN KEY (log_NEXTCOM_email_id) REFERENCES tb_email(email_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ;"
CALL DB.Execute(sql, 67)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 68
'...........................................................................................
'modifica gestione log di spedizione email del nextContract: viene incorporato nel next-com
'...........................................................................................
sql = " ALTER TABLE Ctb_log_mailing DROP CONSTRAINT FK_Ctb_log_mailing__tb_email; " & _
	  " ALTER TABLE Ctb_log_mailing DROP COLUMN log_NEXTCOM_email_id; " & _
	  " ALTER TABLE Ctb_log_mailing ADD log_NEXTCOM_email_id_IT INT NOT NULL, " &_
	  "		log_NEXTCOM_email_id_EN INT NOT NULL ;" & _
	  " ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__tb_email_IT " & _
	  "		FOREIGN KEY (log_NEXTCOM_email_id_IT) REFERENCES tb_email(email_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ; " & _
	  " ALTER TABLE Ctb_log_mailing ADD CONSTRAINT FK_Ctb_log_mailing__tb_email_EN " & _
	  "		FOREIGN KEY (log_NEXTCOM_email_id_EN) REFERENCES tb_email(email_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ; "
CALL DB.Execute(sql, 68)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 69
'...........................................................................................
'modifica gestione log download documenti del nextContract
'...........................................................................................
sql = " ALTER TABLE Ctb_log_download ADD log_lingua_file TEXT(2) WITH COMPRESSION NUlL; " & _
	  " ALTER TABLE Ctb_log_download ADD CONSTRAINT FK_Ctb_log_download__tb_cnt_lingue " & _
	  "		FOREIGN KEY (log_lingua_file) REFERENCES tb_cnt_lingue (lingua_codice) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 69)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 70
'...........................................................................................
'modifica gestione log download documenti del nextContract
'...........................................................................................
sql = " ALTER TABLE Ctb_mail_bandi DROP COLUMN mail_Data "
CALL DB.Execute(sql, 70)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 71
'...........................................................................................
'aggiunge gestione FAQ
'...........................................................................................
sql = "CREATE TABLE tb_FAQ_categorie (" &_
	  "		cat_id COUNTER CONSTRAINT PK_tb_FAQ_categorie PRIMARY KEY, " & _
	  "		cat_nome_IT TEXT(250) WITH COMPRESSION NULL, " & _
	  "		cat_nome_EN TEXT(250) WITH COMPRESSION NULL, " & _
	  "		cat_nome_DE TEXT(250) WITH COMPRESSION NULL, " & _
	  "		cat_nome_FR TEXT(250) WITH COMPRESSION NULL, " & _
	  "		cat_nome_ES TEXT(250) WITH COMPRESSION NULL, " & _
	  "		cat_ordine INT NULL " &_
	  "		);" & _
	  "CREATE TABLE tb_FAQ (" & _
	  "		faq_id COUNTER CONSTRAINT PK_tb_FAQ PRIMARY KEY, " & _
	  "		faq_cat_id INT NOT NULL, " & _ 
	  "		faq_domanda_IT TEXT(250) WITH COMPRESSION NULL, " & _
	  "		faq_domanda_EN TEXT(250) WITH COMPRESSION NULL, " & _
	  "		faq_domanda_DE TEXT(250) WITH COMPRESSION NULL, " & _
	  "		faq_domanda_FR TEXT(250) WITH COMPRESSION NULL, " & _
	  "		faq_domanda_ES TEXT(250) WITH COMPRESSION NULL, " & _
	  "		faq_visibile BIT NULL, " & _
	  "		faq_ordine INT NULL, " & _
	  "		faq_risposta_IT TEXT WITH COMPRESSION NULL, " & _
	  "		faq_risposta_EN TEXT WITH COMPRESSION NULL, " & _
	  "		faq_risposta_DE TEXT WITH COMPRESSION NULL, " & _
	  "		faq_risposta_FR TEXT WITH COMPRESSION NULL, " & _
	  "		faq_risposta_ES TEXT WITH COMPRESSION NULL " & _
	  "		); " & _
	  " ALTER TABLE tb_FAQ ADD CONSTRAINT FK_tb_FAQ__tb_FAQ_categorie " & _
	  "		FOREIGN KEY (faq_cat_id) REFERENCES tb_FAQ_categorie (cat_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 71)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 72
'...........................................................................................
'aggiunge gestione Organigramma
'...........................................................................................
sql = " CREATE TABLE Otb_livelli (" & _
	  "		lvl_id COUNTER CONSTRAINT PK_otb_livelli PRIMARY KEY, " & _
	  "		lvl_nome_IT TEXT(250) WITH COMPRESSION NULL, " & _
	  "		lvl_nome_EN TEXT(250) WITH COMPRESSION NULL, " & _
	  "		lvl_nome_DE TEXT(250) WITH COMPRESSION NULL, " & _
	  "		lvl_nome_FR TEXT(250) WITH COMPRESSION NULL, " & _
	  "		lvl_nome_ES TEXT(250) WITH COMPRESSION NULL, " & _
	  "		lvl_ordine INT NULL " &_
	  "		);" & _
	  " CREATE TABLE Otb_componenti (" & _
	  "		com_id COUNTER CONSTRAINT PK_otb_componenti PRIMARY KEY, " & _
	  "		com_NEXTCOM_id INT NOT NULL, " & _ 
	  "		com_lvl_id INT NOT NULL, " & _ 
	  "		com_visibile BIT NULL, " & _ 
	  "		com_ordine INT NULL, " & _
	  "		com_foto TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_posizione_IT TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_posizione_EN TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_posizione_FR TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_posizione_DE TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_posizione_ES TEXT(250) WITH COMPRESSION NULL, " & _
	  "		com_curriculum_IT TEXT WITH COMPRESSION NULL, " & _
	  "		com_curriculum_EN TEXT WITH COMPRESSION NULL, " & _
	  "		com_curriculum_FR TEXT WITH COMPRESSION NULL, " & _
	  "		com_curriculum_DE TEXT WITH COMPRESSION NULL, " & _
	  "		com_curriculum_ES TEXT WITH COMPRESSION NULL " & _
	  "		); " & _
	  " ALTER TABLE Otb_componenti ADD CONSTRAINT FK_Otb_componenti__Otb_livelli " &_
	  "		FOREIGN KEY (com_lvl_id) REFERENCES Otb_livelli (lvl_id) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ;" & _
	  "	ALTER TABLE Otb_componenti ADD CONSTRAINT FK_Otb_componenti__tb_indirizzario " & _
	  "		FOREIGN KEY (com_NEXTCOM_id) REFERENCES tb_indirizzario(IDElencoIndirizzi) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE ;"
CALL DB.Execute(sql, 72)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 73
'...........................................................................................
'aggiunge gestione parametri di installazione / funzionamento alle applicazioni del next-passport
'...........................................................................................
sql = " CREATE TABLE tb_siti_parametri (" & _
	  "		par_id COUNTER CONSTRAINT PK_tb_siti_parametri PRIMARY KEY, " & _
	  "		par_key TEXT(50) WITH COMPRESSION NOT NULL, " & _
	  "		par_value TEXT(250) WITH COMPRESSION NULL, " & _
	  "		par_sito_id INT NOT NULL " & _
	  "		); " & _
	  "ALTER TABLE tb_siti_parametri ADD CONSTRAINT FK_tb_siti_parametri__tb_siti " & _
	  "		FOREIGN KEY (par_sito_id) REFERENCES tb_siti(id_sito) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 73)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 74
'...........................................................................................
'aggiunge gestione parametri di installazione / funzionamento alle applicazioni del next-passport
'...........................................................................................
sql = "ALTER TABLE tb_attivita ADD COLUMN att_inSospeso BIT NULL"
CALL DB.Execute(sql, 74)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 75
'...........................................................................................
'Corregge nomi applicativi installati nel NEXT-passport
'...........................................................................................
sql = "UPDATE tb_siti SET sito_nome='NEXT-passport [gestione utenti]' WHERE id_sito=" & NEXTPASSPORT & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-web [gestione grafica e contenuti]' WHERE id_sito=" & NEXTWEB & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-com [gestione comunicazioni]' WHERE id_sito=" & NEXTCOM & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-news [gestione news]' WHERE id_sito=" & NEXTNEWS & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-link [gestione link utili]' WHERE id_sito=" & NEXTLINK & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-menu [gestione menu'' e ricette]' WHERE id_sito=" & NEXTMENU & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-flat [gestione appartamenti turistici]' WHERE id_sito=" & NEXTFLAT & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-memo [gestione documenti per area riservata]' WHERE id_sito=" & NEXTMEMO & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-banner [gestione banners pubblicitari]' WHERE id_sito=" & NEXTBANNER & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-club [gestione associati]' WHERE id_sito=" & NEXTCLUB & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-booking [gestione prenotazioni]' WHERE id_sito=" & NEXTBOOKING & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-guestbook [gestione guestbook]' WHERE id_sito=" & NEXTGUESTBOOK & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-contract [gestione bandi ed appalti] ' WHERE id_sito=" & NEXTCONTRACT & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-f.a.q. [gestione frequently asked questions]' WHERE id_sito=" & NEXTFAQ & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-team [gestione organigramma aziendale]' WHERE id_sito=" & NEXTTEAM & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-booking portal [gestione portale di prenotazione]' WHERE id_sito=" & NEXTBOOKINGPORTALE
CALL DB.Execute(sql, 75)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 76
'...........................................................................................
'aggiunge campi ultima modifica per pratiche e documenti
'...........................................................................................
sql = "ALTER TABLE tb_pratiche ADD COLUMN " & vbCrLf & _
	  "		pra_mod_data DATETIME NULL, " & vbCrLf & _
	  "		pra_mod_utente INTEGER NULL; " & vbCrLf & _
	  "ALTER TABLE tb_documenti ADD COLUMN " & vbCrLf & _
	  "		doc_mod_data DATETIME NULL, " & vbCrLf & _
	  "		doc_mod_utente INTEGER NULL"
CALL DB.Execute(sql, 76)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 77
'...........................................................................................
'aggiunge relazione mancante tra rel_documenti_descrittori e tb_descrittori
'...........................................................................................
sql = "ALTER TABLE rel_documenti_descrittori ADD CONSTRAINT FK_rel_documenti_descrittori__tb_descrittori " + _
   	  "FOREIGN KEY (rdd_descrittore_id) REFERENCES tb_descrittori (descr_id) " + _
	  "ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 77)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 78
'...........................................................................................
'aggiornamento fantasma per la creazione delle directory temporanee per ogni utente
'...........................................................................................
sql = "SELECT * FROM tb_admin"
CALL DB.Execute(sql, 78)
'..................................
if DB.last_update_executed then
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	'crea cartella temporanea comune
	path = Application("IMAGE_PATH") & "temp"
	if not fso.FolderExists(path) then
		'crea cartella temporanea generale
		fso.CreateFolder(path)
	end if
		
	'crea cartella temporanea documenti
	path = path & "\docs"
	if not fso.FolderExists(path) then
		fso.CreateFolder(path)
		
		'crea cartelle per ogni utente
		sql = "SELECT DISTINCT admin_login FROM tb_admin"
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		while not rs.eof
			fso.CreateFolder(path & "\" & rs("admin_login"))	
			rs.movenext
		wend
		rs.close
	end if
	
	'rimuove tutte le cartelle temporanee dalle cartelle <AZ_ID>
	set folder = fso.GetFolder(Application("IMAGE_PATH"))
	for each SubFolder in folder.SubFolders
		if isNumeric(SubFolder.name) then
			path = Application("IMAGE_PATH") & SubFolder.name & "\temp"
			if fso.FolderExists(path) then
				fso.DeleteFolder(path)
			end if
		end if
	next
	set fso = nothing
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 79
'...........................................................................................
'aggiunge tabelle per la gestione dei files
'...........................................................................................
sql = " CREATE TABLE tb_Files (" & _
	  "		F_id COUNTER CONSTRAINT PK_tb_files PRIMARY KEY, " & _
	  "		F_original_name TEXT(250) WITH COMPRESSION NULL, " & _
	  "		F_encoded_name TEXT(250) WITH COMPRESSION NULL, " & _
	  "		F_size INT NULL, " & _
	  "		F_Data DATETIME NULL, " & _
	  "		F_base_path TEXT(250) WITH COMPRESSION NULL, " & _
	  "		F_allegato BIT NULL " &_
	  "		);" & _
	  " CREATE TABLE rel_documenti_files(" & _
	  "		rel_id COUNTER CONSTRAINT PK_rel_documenti_files PRIMARY KEY, " & _
	  "		rel_documento_id INT NOT NULL, " & _
	  " 	rel_files_id INT NOT NULL " & _
	  "		);" & _
	  " ALTER TABLE rel_documenti_files ADD CONSTRAINT FK_rel_documenti_files__tb_files " + _
   	  " 	FOREIGN KEY (rel_files_id) REFERENCES tb_files (F_id) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  " ALTER TABLE rel_documenti_files ADD CONSTRAINT FK_rel_documenti_files__tb_documenti " + _
	  "		FOREIGN KEY (rel_documento_id) REFERENCES tb_documenti (doc_id) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 79)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 80
'...........................................................................................
'elimina e ricrea tabella log_cnt_email per ricostruire gli indici
'...........................................................................................
sql = " CREATE TABLE tmp_log_cnt_email(" + _
	  " 	log_id COUNTER CONSTRAINT PK_tmp_log_cnt_email PRIMARY KEY, " + _
	  " 	log_cnt_id INT NOT NULL, " + _
	  " 	log_email_id INT NOT NULL, " + _
	  " 	log_email TEXT(50) WITH COMPRESSION NULL " + _
	  " ); " + _
	  " DROP INDEX log_email_id ON log_cnt_email; " + _
	  " DROP INDEX log_str_id ON log_cnt_email; " + _
	  " DROP INDEX email_log_id ON log_cnt_email; " + _
	  " ALTER TABLE log_cnt_email DROP CONSTRAINT [{1829BC39-660E-409C-8691-928E0B643489}]; " + _
	  " ALTER TABLE log_cnt_email DROP CONSTRAINT [{33D720EC-640D-4330-A9E9-D7045DBCD5C9}]; " + _
	  " INSERT INTO tmp_log_cnt_email (log_cnt_id, log_email_id, log_email) " + _
	  "		SELECT log_cnt_id, log_email_id, log_email FROM log_cnt_email ;" + _
	  " DROP TABLE log_cnt_email;" + _
	  " CREATE TABLE log_cnt_email (" + _
	  "		log_id COUNTER CONSTRAINT PK_log_cnt_email PRIMARY KEY, " + _
	  "		log_cnt_id INT NOT NULL, " + _
	  "		log_email_id INT NOT NULL, " + _
	  "		log_email TEXT(50) WITH COMPRESSION NULL " + _
	  ");" + _
	  " INSERT INTO log_cnt_email (log_cnt_id, log_email_id, log_email) " + _
	  "		SELECT log_cnt_id, log_email_id, log_email FROM tmp_log_cnt_email; " + _
	  " DROP TABLE tmp_log_cnt_email; " + _
	  " ALTER TABLE log_cnt_email ADD CONSTRAINT FK_log_cnt_email__tb_indirizzario " + _
	  "		FOREIGN KEY (log_cnt_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " + _
	  " ALTER TABLE log_cnt_email ADD CONSTRAINT FK_log_cnt_email__tb_email " + _
	  "		FOREIGN KEY (log_email_id) REFERENCES tb_email (email_id) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 80)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 81
'...........................................................................................
'aggiunge campi a tabella files e crea directory docs
'...........................................................................................
sql = " ALTER TABLE tb_files ADD COLUMN " + _
	  "		F_original_path TEXT(250) WITH COMPRESSION NULL, " & _
	  " 	F_encoded_path TEXT(250) WITH COMPRESSION NULL, " & _
	  "		F_LastUpdate DATETIME NULL; " & _
	  " ALTER TABLE tb_files DROP COLUMN F_base_path "
CALL DB.Execute(sql, 81)
'..................................
if DB.last_update_executed then
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	'crea destinazione documenti
	path = Application("IMAGE_PATH") & "docs"
	if not fso.FolderExists(path) then
		'crea cartella generale documenti
		fso.CreateFolder(path)
	end if
	set fso = nothing
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 82
'...........................................................................................
'aggiunge campo per tracciatura utente che chiude l'attivita'
'...........................................................................................
sql = " ALTER TABLE tb_attivita ADD COLUMN " + _
	  "		att_utente_chiusura INT NULL "
CALL DB.Execute(sql, 82)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 83
'...........................................................................................
'elimina vecchia gestione documenti
'...........................................................................................
sql = " ALTER TABLE tb_documenti DROP COLUMN doc_path; " + _
	  " DELETE FROM tb_files; "
CALL DB.Execute(sql, 83)
'..................................
if DB.last_update_executed then
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	'cancella vecchie cartelle pratiche e documenti
	set folder = fso.GetFolder(Application("IMAGE_PATH"))
	for each SubFolder in folder.SubFolders
		if isNumeric(SubFolder.name) then
			path = SubFolder.path & "\docs"
			if fso.FolderExists(path) then
				CALL DeleteNotNumericFolders(fso, path)
			end if
		end if
	next
	set fso = nothing
end if

sub DeleteNotNumericFolders(fso, path)
	dim folder, subfolder
	set folder = fso.GetFolder(path)
	for each SubFolder in folder.SubFolders
		if not isNumeric(SubFolder.name) then
			SubFolder.Delete()
		end if
	next
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 84
'...........................................................................................
'svuota tabelle documenti e files e corregge problema su relazione tb_allegati e tb_documenti
'...........................................................................................
sql = " ALTER TABLE tb_allegati DROP CONSTRAINT FK_tb_allegati__tb_documenti; " + _
	  " ALTER TABLE tb_allegati ADD CONSTRAINT FK_tb_allegati__tb_documenti " + _
   	  " FOREIGN KEY (all_documento_id) REFERENCES tb_documenti (doc_id) " + _
	  " ON UPDATE CASCADE ON DELETE CASCADE; " + _
	  " DELETE FROM tb_files; " & _
	  " DELETE FROM tb_documenti; "
CALL DB.Execute(sql, 84)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 85
'...........................................................................................
'aumenta lunghezza campo nome del sito
'...........................................................................................
sql = " ALTER TABLE tb_siti ALTER COLUMN sito_nome TEXT(250) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 85)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 86
'...........................................................................................
'aggiunge struttura dati per NEXT-realestate
'...........................................................................................
sql = "CREATE TABLE Rtb_strutture ("& _
	  "		st_ID COUNTER CONSTRAINT PK_Rtb_strutture PRIMARY KEY, "& _
	  "		st_denominazione VARCHAR(255) NULL, "& _
	  "		st_descrizione_IT TEXT WITH COMPRESSION NULL, "& _
	  "		st_descrizione_EN TEXT WITH COMPRESSION NULL, "& _
	  "		st_descrizione_FR TEXT WITH COMPRESSION NULL, "& _
	  "		st_descrizione_DE TEXT WITH COMPRESSION NULL, "& _
	  "		st_descrizione_ES TEXT WITH COMPRESSION NULL, "& _
	  "		st_metratura VARCHAR(255) NULL, "& _
	  "		st_prezzo CURRENCY NULL, "& _
	  "		st_ordine INTEGER NULL, "& _
	  "		st_home BIT NULL, "& _
	  "		st_NEXTweb_ps_mappa_location INTEGER NULL, "& _
	  "		st_NEXTweb_ps_info INTEGER NULL, "& _
	  "		st_NEXTweb_ps_mappa_catastale INTEGER NULL, "& _
	  "		st_tipologia_id INTEGER NULL, "& _
	  "		st_categoria_id INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_tipologie ("& _
	  "		ti_ID COUNTER CONSTRAINT PK_Rtb_tipologie PRIMARY KEY, "& _
	  "		ti_nome_IT VARCHAR(100) NULL, "& _
	  "		ti_nome_EN VARCHAR(100) NULL, "& _
	  "		ti_nome_FR VARCHAR(100) NULL, "& _
	  "		ti_nome_DE VARCHAR(100) NULL, "& _
	  "		ti_nome_ES VARCHAR(100) NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_categorie ("& _
	  "		ca_ID COUNTER CONSTRAINT PK_Rtb_categorie PRIMARY KEY, "& _
	  "		ca_nome_IT VARCHAR(100) NULL, "& _
	  "		ca_nome_EN VARCHAR(100) NULL, "& _
	  "		ca_nome_FR VARCHAR(100) NULL, "& _
	  "		ca_nome_DE VARCHAR(100) NULL, "& _
	  "		ca_nome_ES VARCHAR(100) NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_caratteristiche ("& _
	  "		car_ID COUNTER CONSTRAINT PK_Rtb_caratteristiche PRIMARY KEY, "& _
	  "		car_nome_IT VARCHAR(255) NULL, "& _
	  "		car_nome_EN VARCHAR(255) NULL, "& _
	  "		car_nome_FR VARCHAR(255) NULL, "& _
	  "		car_nome_DE VARCHAR(255) NULL, "& _
	  "		car_nome_ES VARCHAR(255) NULL, "& _
	  "		car_tipo INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_strutture_caratteristiche ("& _
	  "		sc_ID COUNTER CONSTRAINT PK_Rtb_strutture_caratteristiche PRIMARY KEY, "& _
	  "		sc_valore_IT VARCHAR(255) NULL, "& _
	  "		sc_valore_EN VARCHAR(255) NULL, "& _
	  "		sc_valore_FR VARCHAR(255) NULL, "& _
	  "		sc_valore_ES VARCHAR(255) NULL, "& _
	  "		sc_valore_DE VARCHAR(255) NULL, "& _
	  "		sc_struttura_id INTEGER NULL, "& _
	  "		sc_caratteristica_id INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_foto ("& _
	  "		fo_ID COUNTER CONSTRAINT PK_Rtb_foto PRIMARY KEY, "& _
	  "		fo_image VARCHAR(255) NULL, "& _
	  "		fo_image_zoom VARCHAR(255) NULL, "& _
	  "		fo_didascalia_IT TEXT WITH COMPRESSION NULL, "& _
	  "		fo_didascalia_EN TEXT WITH COMPRESSION NULL, "& _
	  "		fo_didascalia_FR TEXT WITH COMPRESSION NULL, "& _
	  "		fo_didascalia_DE TEXT WITH COMPRESSION NULL, "& _
	  "		fo_didascalia_ES TEXT WITH COMPRESSION NULL, "& _
	  "		fo_ordine INTEGER NULL, "& _
	  "		fo_struttura_id INTEGER NULL"& _
	  ");"& _
	  "CREATE TABLE Rtb_richieste_info ("& _
	  "		ri_ID COUNTER CONSTRAINT PK_Rtb_richieste_info PRIMARY KEY, "& _
	  "		ri_prezzo CURRENCY NULL, "& _
	  "		ri_richiesta TEXT WITH COMPRESSION NULL, "& _
	  "		ri_NEXTcom_ID INTEGER NULL, "& _
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
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 86)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 87
'...........................................................................................
'aggiunge altri campi per NEXT-realestate
'...........................................................................................
sql = "ALTER TABLE Rtb_caratteristiche ADD COLUMN "& _
	  "		car_ordine INTEGER NULL, "& _
	  "		car_icona VARCHAR(255) NULL"
CALL DB.Execute(sql, 87)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 88
'...........................................................................................
'ancora campi per NEXT-realestate
'...........................................................................................
sql = "ALTER TABLE Rtb_richieste_info ADD COLUMN "& _
	  "		ri_data DATETIME NULL"
CALL DB.Execute(sql, 88)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 89
'...........................................................................................
'ancora campi per NEXT-realestate
'...........................................................................................
sql = "ALTER TABLE Rtb_categorie ADD COLUMN "& _
	  "		ca_ordine INTEGER NULL, "& _
	  "		ca_icona VARCHAR(255) NULL"
CALL DB.Execute(sql, 89)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 90
'...........................................................................................
'tabella citta per NEXT-realestate
'...........................................................................................
sql = "CREATE TABLE Rtb_citta ("& _
	  "		ci_ID COUNTER CONSTRAINT PK_Rtb_citta PRIMARY KEY, "& _
	  "		ci_nome_IT VARCHAR(255) NULL, "& _
	  "		ci_nome_EN VARCHAR(255) NULL, "& _
	  "		ci_nome_FR VARCHAR(255) NULL, "& _
	  "		ci_nome_ES VARCHAR(255) NULL, "& _
	  "		ci_nome_DE VARCHAR(255) NULL"& _
	  ");"& _
	  "ALTER TABLE Rtb_strutture ADD COLUMN "& _
	  "		st_citta_id INTEGER NULL;"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_citta "& _
   	  "		FOREIGN KEY (st_citta_id) REFERENCES Rtb_citta (ci_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 90)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 91
'...........................................................................................
'il prezzo in txt
'...........................................................................................
sql = "ALTER TABLE Rtb_strutture ALTER COLUMN "& _
	  "		st_prezzo VARCHAR(255) NULL"
CALL DB.Execute(sql, 91)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 92
'...........................................................................................
'prezzo e denominazione multilingua
'...........................................................................................
sql = "ALTER TABLE Rtb_strutture DROP COLUMN st_prezzo;"& _
	  "ALTER TABLE Rtb_strutture DROP COLUMN st_denominazione;"& _
	  "ALTER TABLE Rtb_strutture ADD COLUMN "& _
	  "		st_denominazione_it VARCHAR(255) NULL, "& _
	  "		st_denominazione_en VARCHAR(255) NULL, "& _
	  "		st_denominazione_fr VARCHAR(255) NULL, "& _
	  "		st_denominazione_es VARCHAR(255) NULL, "& _
	  "		st_denominazione_de VARCHAR(255) NULL, "& _
	  "		st_prezzo_it VARCHAR(255) NULL, "& _
	  "		st_prezzo_en VARCHAR(255) NULL, "& _
	  "		st_prezzo_fr VARCHAR(255) NULL, "& _
	  "		st_prezzo_es VARCHAR(255) NULL, "& _
	  "		st_prezzo_de VARCHAR(255) NULL"
CALL DB.Execute(sql, 92)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 93
'...........................................................................................
'metratura multilingua
'...........................................................................................
sql = "ALTER TABLE Rtb_strutture DROP COLUMN st_metratura;"& _
	  "ALTER TABLE Rtb_strutture ADD COLUMN "& _
	  "		st_metratura_it VARCHAR(255) NULL, "& _
	  "		st_metratura_en VARCHAR(255) NULL, "& _
	  "		st_metratura_fr VARCHAR(255) NULL, "& _
	  "		st_metratura_es VARCHAR(255) NULL, "& _
	  "		st_metratura_de VARCHAR(255) NULL"
CALL DB.Execute(sql, 93)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 94
'...........................................................................................
'aggiornamento per modifica permessi di accesso degli utenti su area riservata della parte
'visibile: cambiata procedura CheckLogin, ora verifica anche abilitazione e data scadenza
'...........................................................................................
sql = "UPDATE tb_utenti SET ut_abilitato=true"
CALL DB.Execute(sql, 94)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 95
'...........................................................................................
'tabella contratti per NEXT-realestate
'...........................................................................................
sql = "CREATE TABLE Rtb_contratti ("& _
	  "		co_ID COUNTER CONSTRAINT PK_Rtb_contratti PRIMARY KEY, "& _
	  "		co_nome_IT VARCHAR(255) NULL, "& _
	  "		co_nome_EN VARCHAR(255) NULL, "& _
	  "		co_nome_FR VARCHAR(255) NULL, "& _
	  "		co_nome_ES VARCHAR(255) NULL, "& _
	  "		co_nome_DE VARCHAR(255) NULL"& _
	  ");"& _
	  "ALTER TABLE Rtb_strutture ADD COLUMN "& _
	  "		st_contratto_id INTEGER NULL;"& _
	  "ALTER TABLE Rtb_strutture ADD CONSTRAINT FK_Rtb_strutture__Rtb_contratti "& _
   	  "		FOREIGN KEY (st_contratto_id) REFERENCES Rtb_contratti (co_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 95)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 96
'...........................................................................................
'flag visibile per strutture NEXT-realestate
'...........................................................................................
sql = "ALTER TABLE Rtb_strutture ADD COLUMN "& _
	  "		st_visibile BIT NULL; "& _
	  "UPDATE Rtb_strutture SET st_visibile=1"
CALL DB.Execute(sql, 96)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 97
'...........................................................................................
'flag visibile per strutture NEXT-flat
'...........................................................................................
sql = "ALTER TABLE Atb_appartamenti ADD COLUMN "& _
	  "		ap_abilitato BIT NULL; "& _
	  "UPDATE Atb_appartamenti SET ap_abilitato=1"
CALL DB.Execute(sql, 97)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 98
'...........................................................................................
'ripulisce database da prenotazioni non valide
'...........................................................................................
sql = "DELETE FROM btb_prenotazioni WHERE pre_cliente_id NOT IN (SELECT IDElencoIndirizzi FROM tb_indirizzario) "
CALL DB.Execute(sql, 98)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 99
'...........................................................................................
'aggiunge relazione per NEXTbooking tra prenotazione e contatto.
'...........................................................................................
sql = "ALTER TABLE btb_prenotazioni ADD CONSTRAINT FK_btb_prenotazioni__tb_indirizzario " & _
   	  "		FOREIGN KEY (pre_cliente_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) " & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 99)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 100
'...........................................................................................
'aumenta dimensione campo email su log di spedizione
'...........................................................................................
sql = "ALTER TABLE log_cnt_email ALTER COLUMN log_email nvarchar(250) "
CALL DB.Execute(sql, 100)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 101
'...........................................................................................
'aggiunge tabelle e relazioni per l'applicativo verticale TimeTable (vedi Favret)
'...........................................................................................
sql = "CREATE TABLE Ftb_localita ("& _
	  "		lo_ID COUNTER CONSTRAINT PK_Ftb_localita PRIMARY KEY, "& _
	  "		lo_nome_it VARCHAR(100) NULL, "& _
	  "		lo_nome_en VARCHAR(100) NULL, "& _
	  "		lo_nome_fr VARCHAR(100) NULL, "& _
	  "		lo_nome_es VARCHAR(100) NULL, "& _
	  "		lo_nome_de VARCHAR(100) NULL"& _
	  ");"& _
	  "CREATE TABLE Ftb_tratte ("& _
	  "		tr_ID COUNTER CONSTRAINT PK_Ftb_tratte PRIMARY KEY, "& _
	  "		tr_da_id INTEGER NULL, "& _
	  "		tr_a_id INTEGER NULL, "& _
	  "		tr_ritorno BIT NULL, "& _
	  "		tr_visibile BIT NULL, "& _
	  "		tr_ordine INTEGER NULL, "& _
	  "		tr_note_it TEXT WITH COMPRESSION NULL, "& _
	  "		tr_note_en TEXT WITH COMPRESSION NULL, "& _
	  "		tr_note_fr TEXT WITH COMPRESSION NULL, "& _
	  "		tr_note_es TEXT WITH COMPRESSION NULL, "& _
	  "		tr_note_de TEXT WITH COMPRESSION NULL"& _
	  ");"& _
	  "ALTER TABLE Ftb_tratte ADD CONSTRAINT FK_Ftb_tratte__Ftb_localita "& _
   	  "		FOREIGN KEY (tr_da_id) REFERENCES Ftb_localita (lo_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "CREATE TABLE Ftb_periodi ("& _
	  "		pe_ID COUNTER CONSTRAINT PK_Ftb_periodi PRIMARY KEY, "& _
	  "		pe_dal DATETIME NULL, "& _
	  "		pe_al DATETIME NULL"& _
	  ");"& _
	  "CREATE TABLE Ftb_giorniS ("& _
	  "		gs_ID COUNTER CONSTRAINT PK_Ftb_giorniS PRIMARY KEY, "& _
	  "		gs_giorno INTEGER NULL, "& _
	  "		gs_partenzaA VARCHAR(50) NULL, "& _
	  "		gs_arrivoA VARCHAR(50) NULL, "& _
	  "		gs_partenzaR VARCHAR(50) NULL, "& _
	  "		gs_arrivoR VARCHAR(50) NULL, "& _
	  "		gs_scalo BIT NULL, "& _
	  "		gs_note_it TEXT WITH COMPRESSION NULL, "& _
	  "		gs_note_en TEXT WITH COMPRESSION NULL, "& _
	  "		gs_note_fr TEXT WITH COMPRESSION NULL, "& _
	  "		gs_note_es TEXT WITH COMPRESSION NULL, "& _
	  "		gs_note_de TEXT WITH COMPRESSION NULL, "& _
	  "		gs_periodo_id INTEGER NULL, "& _
	  "		gs_tratta_id INTEGER NULL"& _
	  ");"& _
	  "ALTER TABLE Ftb_giorniS ADD CONSTRAINT FK_Ftb_giorniS__Ftb_periodi "& _
   	  "		FOREIGN KEY (gs_periodo_id) REFERENCES Ftb_periodi (pe_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Ftb_giorniS ADD CONSTRAINT FK_Ftb_giorniS__Ftb_tratte "& _
   	  "		FOREIGN KEY (gs_tratta_id) REFERENCES Ftb_tratte (tr_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 101)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 102
'...........................................................................................
'aggiunta campi nel guestbook NEXT-guestbook
'...........................................................................................
sql = "ALTER TABLE tb_guestbook ADD COLUMN " + _
	  "		Oggetto TEXT(250) WITH COMPRESSION NULL, " + _
	  "		Log_richiesta TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 102)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 103
'...........................................................................................
'aggiunta campo nave su TimeTable
'...........................................................................................
sql = "ALTER TABLE ftb_tratte ADD COLUMN " + _
	  "		tr_nave VARCHAR(250) NULL"
CALL DB.Execute(sql, 103)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 104
'...........................................................................................
'aggiunta campo scalo su TimeTable
'...........................................................................................
sql = "ALTER TABLE ftb_giorniS ADD COLUMN " + _
	  "		gs_scaloA BIT NULL,"& _
	  "		gs_scaloR BIT NULL;"& _
	  "ALTER TABLE ftb_giorniS DROP COLUMN "& _
	  "		gs_scalo"
CALL DB.Execute(sql, 104)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 105
'...........................................................................................
'aggiunge tabelle e relazioni per NEXTschool
'...........................................................................................
sql = "CREATE TABLE Stb_docenti ("& _
	  "		do_ID INTEGER CONSTRAINT PK_Stb_docenti PRIMARY KEY, "& _
	  "		do_classeConcorso VARCHAR(50) NULL, "& _
	  "		do_laurea_it VARCHAR(250) NULL, "& _
	  "		do_laurea_en VARCHAR(250) NULL, "& _
	  "		do_laurea_fr VARCHAR(250) NULL, "& _
	  "		do_laurea_es VARCHAR(250) NULL, "& _
	  "		do_laurea_de VARCHAR(250) NULL, "& _
	  "		do_ricevimento_it VARCHAR(250) NULL, "& _
	  "		do_ricevimento_en VARCHAR(250) NULL, "& _
	  "		do_ricevimento_fr VARCHAR(250) NULL, "& _
	  "		do_ricevimento_es VARCHAR(250) NULL, "& _
	  "		do_ricevimento_de VARCHAR(250) NULL "& _
	  ");"& _
	  "ALTER TABLE Stb_docenti ADD CONSTRAINT FK_Stb_docenti__tb_indirizzario "& _
   	  "		FOREIGN KEY (do_id) REFERENCES tb_indirizzario (IDElencoIndirizzi) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "CREATE TABLE Stb_materie ("& _
	  "		ma_ID COUNTER CONSTRAINT PK_Stb_materie PRIMARY KEY, "& _
	  "		ma_classeConcorso VARCHAR(50) NULL, "& _
	  "		ma_nome_it VARCHAR(250) NULL, "& _
	  "		ma_nome_en VARCHAR(250) NULL, "& _
	  "		ma_nome_fr VARCHAR(250) NULL, "& _
	  "		ma_nome_es VARCHAR(250) NULL, "& _
	  "		ma_nome_de VARCHAR(250) NULL, "& _
	  "		ma_obiettivi_it TEXT WITH COMPRESSION NULL, "& _
	  "		ma_obiettivi_en TEXT WITH COMPRESSION NULL, "& _
	  "		ma_obiettivi_fr TEXT WITH COMPRESSION NULL, "& _
	  "		ma_obiettivi_es TEXT WITH COMPRESSION NULL, "& _
	  "		ma_obiettivi_de TEXT WITH COMPRESSION NULL, "& _
	  "		ma_programma_it TEXT WITH COMPRESSION NULL, "& _
	  "		ma_programma_en TEXT WITH COMPRESSION NULL, "& _
	  "		ma_programma_fr TEXT WITH COMPRESSION NULL, "& _
	  "		ma_programma_es TEXT WITH COMPRESSION NULL, "& _
	  "		ma_programma_de TEXT WITH COMPRESSION NULL, "& _
	  "		ma_esame_it VARCHAR(100) NULL, "& _
	  "		ma_esame_en VARCHAR(100) NULL, "& _
	  "		ma_esame_fr VARCHAR(100) NULL, "& _
	  "		ma_esame_es VARCHAR(100) NULL, "& _
	  "		ma_esame_de VARCHAR(100) NULL "& _
	  ");"& _
	  "CREATE TABLE Stb_classi ("& _
	  "		cl_ID COUNTER CONSTRAINT PK_Stb_classi PRIMARY KEY, "& _
	  "		cl_numero INTEGER NULL, "& _
	  "		cl_sezione VARCHAR(50) NULL, "& _
	  "		cl_anno INTEGER NULL, "& _
	  "		cl_testi VARCHAR(250) NULL, "& _
	  "		cl_testiFile VARCHAR(255) NULL, "& _
	  "		cl_note TEXT WITH COMPRESSION NULL "& _
	  ");"& _
	  "CREATE TABLE Srel_insegna ("& _
	  "		in_ID COUNTER CONSTRAINT PK_Srel_insegna PRIMARY KEY, "& _
	  "		in_docente_id INTEGER NULL, "& _
	  "		in_materia_id INTEGER NULL, "& _
	  "		in_classe_id INTEGER NULL "& _
	  ");"& _
	  "ALTER TABLE Srel_insegna ADD CONSTRAINT FK_Srel_insegna__Stb_docenti "& _
   	  "		FOREIGN KEY (in_docente_id) REFERENCES Stb_docenti (do_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Srel_insegna ADD CONSTRAINT FK_Srel_insegna__Stb_materie "& _
   	  "		FOREIGN KEY (in_materia_id) REFERENCES Stb_materie (ma_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Srel_insegna ADD CONSTRAINT FK_Srel_insegna__Stb_classi "& _
   	  "		FOREIGN KEY (in_classe_id) REFERENCES Stb_classi (cl_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "CREATE TABLE Stb_allievi ("& _
	  "		al_ID COUNTER CONSTRAINT PK_Stb_allievi PRIMARY KEY, "& _
	  "		al_nome VARCHAR(100) NULL, "& _
	  "		al_cognome VARCHAR(100) NULL, "& _
	  "		al_annoNascita INTEGER NULL, "& _
	  "		al_matricola VARCHAR(20) NULL, "& _
	  "		al_note TEXT WITH COMPRESSION NULL "& _
	  ");"& _
	  "CREATE TABLE Srel_studia ("& _
	  "		st_ID COUNTER CONSTRAINT PK_Srel_studia PRIMARY KEY, "& _
	  "		st_allievo_id INTEGER NULL, "& _
	  "		st_classe_id INTEGER NULL "& _
	  ");"& _
	  "ALTER TABLE Srel_studia ADD CONSTRAINT FK_Srel_studia__Stb_allievi "& _
   	  "		FOREIGN KEY (st_allievo_id) REFERENCES Stb_allievi (al_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& _
	  "ALTER TABLE Srel_studia ADD CONSTRAINT FK_Srel_studia__Stb_classi "& _
   	  "		FOREIGN KEY (st_classe_id) REFERENCES Stb_classi (cl_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE"
CALL DB.Execute(sql, 105)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 106
'...........................................................................................
'aggiunge tabelle e relazioni per l'applicativo verticale Proforma di Conto esborsi (vedi Favret)
'...........................................................................................
sql = "CREATE TABLE Ftb_tipi_Nave ("& _
	  "		tpnv_ID COUNTER CONSTRAINT PK_Ftb_tipo_nave PRIMARY KEY, "& _
	  "		tpnv_nome VARCHAR(100) NULL "& _
	  ");"& _
	  "CREATE TABLE Ftb_proforma_esborsi ("& _
	  "		prf_ID COUNTER CONSTRAINT PK_Ftb_proforma_esborsi PRIMARY KEY, "& _
	  "		prf_id_tipo_nave INTEGER NOT NULL, "& _
	  "		prf_nomeNave VARCHAR(100) NULL, "& _
	  "		prf_bandieraNave VARCHAR(100) NULL, "& _
	  "		prf_Dwcc VARCHAR(15) NULL, "& _
	  "		prf_SLT VARCHAR(15) NULL, "& _
	  "		prf_SNT VARCHAR(15) NULL, "& _
	  "		prf_lungh_metri VARCHAR(15) NULL, "& _
	  "		prf_largh_metri VARCHAR(15) NULL, "& _
	  "		prf_Pescaggio_metri VARCHAR(15) NULL, "&  _
	  "		prf_Equipaggio_n VARCHAR(15) NULL, " & _ 	
	  "		prf_Tassa_ancoraggio VARCHAR(15) NULL, " & _
	  "		prf_portoApprodo VARCHAR(100) NULL, "& _
	  "		prf_imbarca_tonn VARCHAR(15) NULL, " & _ 	
	  "		prf_sbarca_tonn VARCHAR(15) NULL, " & _
	  "		prf_descr_merce VARCHAR(200) NULL, " & _  
	  "		prf_tipo_merce INTEGER NULL, "& _
	  "		prf_IMO INTEGER NULL, "& _
	  "		prf_ormeggio_ric VARCHAR(100) NULL, " & _
	  "		prf_provenienza_carico VARCHAR(100) NULL, " & _
	  "		prf_ETA_provenienza_carico VARCHAR(15) NULL, " & _
	  "		prf_destinazione_carico VARCHAR(100) NULL, " & _
	  "		prf_ETA_destinazione_carico VARCHAR(15) NULL, " & _
	  "		prf_id_cliente INTEGER NULL, "& _
	  "		prf_tipo_cliente INTEGER NULL, "& _
	  "		prf_note TEXT WITH COMPRESSION NULL, "& _
	  "		prf_DataRichiesta DATETIME NULL "& _
	  ");"& _
	  "ALTER TABLE Ftb_proforma_esborsi ADD CONSTRAINT FK_Ftb_proforma_esborsi__Ftb_tipi_Nave "& _
   	  "		FOREIGN KEY (prf_id_tipo_nave) REFERENCES Ftb_tipi_Nave (tpnv_ID) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
	CALL DB.Execute(sql, 106)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 107
'...........................................................................................
'aggiunta campo sede nelle classi scuola media trentin
'...........................................................................................
sql = "ALTER TABLE stb_classi ADD COLUMN " + _
	  "cl_sede VARCHAR(200) NULL;"
CALL DB.Execute(sql, 107)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 108
'...........................................................................................
'aggiunge campo per l'applicativo verticale Proforma di Conto esborsi (vedi Favret)
'...........................................................................................
sql = "ALTER TABLE Ftb_proforma_esborsi ADD COLUMN "& _
	  "		prf_riferimento VARCHAR(250) NULL;"
	CALL DB.Execute(sql, 108)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 109
'...........................................................................................
'aggiunge campo per l'applicativo verticale Proforma di Conto esborsi (vedi Favret)
'...........................................................................................
sql = "ALTER TABLE Ftb_tipi_nave DROP COLUMN "& _
	  "		tpnv_nome;"& _
	  "ALTER TABLE Ftb_tipi_nave ADD COLUMN "& _
	  "		tpnv_nome_it VARCHAR(100), "& _
	  "		tpnv_nome_en VARCHAR(100), "& _
	  "		tpnv_nome_fr VARCHAR(100), "& _
	  "		tpnv_nome_es VARCHAR(100), "& _
	  "		tpnv_nome_de VARCHAR(100)"
	CALL DB.Execute(sql, 109)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 110
'...........................................................................................
'aggiunge gestione docenti next-school inserita la classe di concorso
'...........................................................................................
sql = " CREATE TABLE Stb_ClassiConcorso (" & _
	  "		CC_id COUNTER CONSTRAINT PK_Stb_ClassiConcorso PRIMARY KEY, " & _
	  "		CC_nome_IT VARCHAR(250) NULL, " & _
	  "		CC_nome_EN VARCHAR(250) NULL, " & _
	  "		CC_nome_DE VARCHAR(250) NULL, " & _
	  "		CC_nome_ES VARCHAR(250) NULL, " & _
	  "		CC_nome_FR VARCHAR(250) NULL, " & _
	  "		CC_Sigla VARCHAR(10) NULL, " & _
	  "		CC_Sigla_ROM VARCHAR(10) NULL " & _
	  "		);"
CALL DB.Execute(sql, 110)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 111
'...........................................................................................
'aggiunge gestione docenti next-school inserita la classe di concorso
'...........................................................................................
sql = " ALTER TABLE Stb_docenti DROP COLUMN "& _
	  "		do_classeConcorso;"& _
	  " ALTER TABLE Stb_docenti ADD COLUMN "& _
	  "		do_classeConcorso_id INT;" & _
	  " ALTER TABLE Stb_docenti ADD CONSTRAINT FK_Stb_docenti__Stb_ClassiConcorso "& _
   	  "		FOREIGN KEY (do_classeConcorso_id) REFERENCES Stb_ClassiConcorso (CC_id) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;" & _
	  " ALTER TABLE Stb_materie DROP COLUMN "& _
	  "		ma_classeConcorso;"& _
	  " ALTER TABLE Stb_materie ADD COLUMN "& _
	  "		ma_classeConcorso_id INT;" & _
	  " ALTER TABLE Stb_materie ADD CONSTRAINT FK_Stb_materie__Stb_ClassiConcorso "& _
   	  "		FOREIGN KEY (ma_classeConcorso_id) REFERENCES Stb_ClassiConcorso (CC_id) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;" & _
	  " CREATE TABLE Stb_Sedi (" & _
	  "		SD_id COUNTER CONSTRAINT PK_Stb_Sedi PRIMARY KEY, " & _
	  "		SD_nome VARCHAR(250) NULL " & _
	  "		);" & _
	  " ALTER TABLE Stb_classi DROP COLUMN "& _
	  "		cl_sede;"& _
	  " ALTER TABLE Stb_classi ADD COLUMN "& _
	  "		cl_sede_id INT" & _
	  "		);" & _
	  " ALTER TABLE Stb_classi ADD CONSTRAINT FK_Stb_classi__Stb_Sedi "& _
   	  "		FOREIGN KEY (cl_sede_id) REFERENCES Stb_Sedi (SD_id) "& _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 111)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 112
'...........................................................................................
'aggiunge tabelle e relazioni per NEXTtravel
'...........................................................................................
sql = "CREATE TABLE Ttb_categorie ("& vbCrLf & _
	  "		ca_ID COUNTER CONSTRAINT PK_Ttb_categorie PRIMARY KEY, "& vbCrLf & _
	  "		ca_nome_it VARCHAR(150) NULL, "& vbCrLf & _
	  "		ca_nome_en VARCHAR(150) NULL, "& vbCrLf & _
	  "		ca_nome_fr VARCHAR(150) NULL, "& vbCrLf & _
	  "		ca_nome_es VARCHAR(150) NULL, "& vbCrLf & _
	  "		ca_nome_de VARCHAR(150) NULL, "& vbCrLf & _
	  "		ca_file VARCHAR(255) NULL, "& vbCrLf & _
	  "		ca_ordine INTEGER NULL, "& vbCrLf & _
	  "		ca_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		ca_descr_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		ca_descr_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		ca_descr_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		ca_descr_de TEXT WITH COMPRESSION NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE Ttb_sottoCategorie ("& vbCrLf & _
	  "		sc_ID COUNTER CONSTRAINT PK_Ttb_sottoCategorie PRIMARY KEY, "& vbCrLf & _
	  "		sc_categoria_id INTEGER NULL, "& vbCrLf & _
	  "		sc_nome_it VARCHAR(150) NULL, "& vbCrLf & _
	  "		sc_nome_en VARCHAR(150) NULL, "& vbCrLf & _
	  "		sc_nome_fr VARCHAR(150) NULL, "& vbCrLf & _
	  "		sc_nome_es VARCHAR(150) NULL, "& vbCrLf & _
	  "		sc_nome_de VARCHAR(150) NULL, "& vbCrLf & _
	  "		sc_file VARCHAR(255) NULL, "& vbCrLf & _
	  "		sc_link VARCHAR(255) NULL, "& vbCrLf & _
	  "		sc_ordine INTEGER NULL, "& vbCrLf & _
	  "		sc_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		sc_descr_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		sc_descr_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		sc_descr_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		sc_descr_de TEXT WITH COMPRESSION NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Ttb_sottoCategorie ADD CONSTRAINT FK_Ttb_sottoCategorie__Ttb_categorie "& vbCrLf & _
   	  "		FOREIGN KEY (sc_categoria_id) REFERENCES Ttb_categorie (ca_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "CREATE TABLE Ttb_destinazioni ("& vbCrLf & _
	  "		de_ID COUNTER CONSTRAINT PK_Ttb_destinazioni PRIMARY KEY, "& vbCrLf & _
   	  "		de_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_file VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_link VARCHAR(255) NULL, "& vbCrLf & _
	  "		de_ordine INTEGER NULL, "& vbCrLf & _
	  "		de_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		de_descr_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		de_descr_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		de_descr_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		de_descr_de TEXT WITH COMPRESSION NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE Ttb_viaggi ("& vbCrLf & _
	  "		vi_ID COUNTER CONSTRAINT PK_Ttb_viaggi PRIMARY KEY, "& vbCrLf & _
	  "		vi_destinazione_id INTEGER NULL, "& vbCrLf & _
   	  "		vi_nome_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_nome_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_nome_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_nome_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_nome_de TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_ordine INTEGER NULL, "& vbCrLf & _
	  "		vi_visibile BIT NULL, "& vbCrLf & _
	  "		vi_NEXTweb_ps INTEGER NULL, "& vbCrLf & _
	  "		vi_file VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_partenza VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_notti VARCHAR(100) NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ADD CONSTRAINT FK_Ttb_viaggi__Ttb_destinazioni "& vbCrLf & _
   	  "		FOREIGN KEY (vi_destinazione_id) REFERENCES Ttb_destinazioni (de_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "CREATE TABLE Trel_viaggi_sottoCategorie ("& vbCrLf & _
	  "		rvs_ID COUNTER CONSTRAINT PK_Trel_viaggi_sottoCategorie PRIMARY KEY, "& vbCrLf & _
	  "		rvs_viaggio_id INTEGER NULL, "& vbCrLf & _
	  "		rvs_sottoCategoria_id INTEGER NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Trel_viaggi_sottoCategorie ADD CONSTRAINT FK_Trel_viaggi_sottoCategorie__Ttb_viaggi "& vbCrLf & _
   	  "		FOREIGN KEY (rvs_viaggio_id) REFERENCES Ttb_viaggi (vi_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE Trel_viaggi_sottoCategorie ADD CONSTRAINT FK_Trel_viaggi_sottoCategorie__Ttb_sottoCategorie "& vbCrLf & _
   	  "		FOREIGN KEY (rvs_sottoCategoria_id) REFERENCES Ttb_sottoCategorie (sc_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "CREATE TABLE Trel_info ("& vbCrLf & _
	  "		in_ID COUNTER CONSTRAINT PK_Trel_info PRIMARY KEY, "& vbCrLf & _
	  "		in_viaggio_id INTEGER NULL, "& vbCrLf & _
	  "		in_indirizzario_id INTEGER NULL, "& vbCrLf & _
	  "		in_data DATETIME NULL, "& vbCrLf & _
   	  "		in_info TEXT WITH COMPRESSION NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Trel_info ADD CONSTRAINT FK_Trel_info__Ttb_viaggi "& vbCrLf & _
   	  "		FOREIGN KEY (in_viaggio_id) REFERENCES Ttb_viaggi (vi_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE Trel_info ADD CONSTRAINT FK_Trel_info__tb_indirizzario "& vbCrLf & _
   	  "		FOREIGN KEY (in_indirizzario_id) REFERENCES tb_indirizzario (idElencoIndirizzi) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "CREATE TABLE Ttb_descrittori ("& vbCrLf & _
	  "		des_ID COUNTER CONSTRAINT PK_Ttb_descrittori PRIMARY KEY, "& vbCrLf & _
	  "		des_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_tipo INTEGER NULL, "& vbCrLf & _
	  "		des_ordine INTEGER NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE Trel_descrittori_viaggi ("& vbCrLf & _
	  "		rdv_ID COUNTER CONSTRAINT PK_Trel_descrittori_viaggi PRIMARY KEY, "& vbCrLf & _
	  "		rdv_descrittore_id INTEGER NULL, "& vbCrLf & _
	  "		rdv_viaggio_id INTEGER NULL, "& vbCrLf & _
	  "		rdv_valore_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdv_valore_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdv_valore_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdv_valore_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdv_valore_de VARCHAR(255) NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Trel_descrittori_viaggi ADD CONSTRAINT FK_Trel_descrittori_viaggi__Ttb_descrittori "& vbCrLf & _
   	  "		FOREIGN KEY (rdv_descrittore_id) REFERENCES Ttb_descrittori (des_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE Trel_descrittori_viaggi ADD CONSTRAINT FK_Trel_descrittori_viaggi__Ttb_viaggi "& vbCrLf & _
   	  "		FOREIGN KEY (rdv_viaggio_id) REFERENCES Ttb_viaggi (vi_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 112)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 113
'...........................................................................................
'aggiunge campi a tabella viaggi per NEXTtravel
'...........................................................................................
sql = "ALTER TABLE Ttb_viaggi ALTER COLUMN vi_nome_it VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ALTER COLUMN vi_nome_en VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ALTER COLUMN vi_nome_fr VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ALTER COLUMN vi_nome_es VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ALTER COLUMN vi_nome_de VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi ADD COLUMN "& vbCrLf & _
	  "		vi_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_descr_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_descr_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_descr_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		vi_descr_de TEXT WITH COMPRESSION NULL;"
CALL DB.Execute(sql, 113)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 114
'...........................................................................................
'aggiunge campi a tabella dei listini del next booking
'...........................................................................................
sql = "ALTER TABLE btb_listini ADD COLUMN "& vbCrLf & _
	  "		lis_condizioni_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		lis_condizioni_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		lis_condizioni_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		lis_condizioni_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		lis_condizioni_de TEXT WITH COMPRESSION NULL;"
CALL DB.Execute(sql, 114)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 115
'...........................................................................................
'aggiunge campi a tabella dei tipi di camera del next booking
'...........................................................................................
sql = "ALTER TABLE btb_tipiCamera ADD COLUMN "& vbCrLf & _
	 "		tipC_immagine TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  "		tipC_descrizione_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		tipC_descrizione_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		tipC_descrizione_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		tipC_descrizione_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		tipC_descrizione_de TEXT WITH COMPRESSION NULL;"
CALL DB.Execute(sql, 115)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 116
'...........................................................................................
'aggiunge campi a tabella prenotazioni del next booking
'...........................................................................................
sql = "ALTER TABLE btb_prenotazioni ADD COLUMN "& vbCrLf & _
	 "		pre_meseCC int NULL, " & vbCrLf & _
	  "		pre_annoCC int NULL;"
CALL DB.Execute(sql, 116)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 117
'...........................................................................................
'aggiunge campi a tabella prenotazioni del next booking
'...........................................................................................
sql = "ALTER TABLE btb_prenotazioni ADD COLUMN "& vbCrLf & _
	 "		pre_cvcCC TEXT(5) NULL;"
CALL DB.ProtectedExecute(sql, 117, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 118
'...........................................................................................
'aggiunge campi e tabella sezioni per NEXTtravel
'...........................................................................................
sql = "CREATE TABLE Ttb_sezioni ("& vbCrLf & _
	  "		se_ID COUNTER CONSTRAINT PK_Ttb_sezioni PRIMARY KEY, "& vbCrLf & _
	  "		se_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_title_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_title_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_title_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_title_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_title_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		se_descrizione_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		se_descrizione_en TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		se_descrizione_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		se_descrizione_es TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		se_descrizione_de TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		se_visibile BIT NULL, "& vbCrLf & _
	  "		se_ordine INTEGER NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE Ttb_categorie ADD COLUMN "& vbCrLf & _
	  "		ca_sezione_id INTEGER NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_categorie ADD CONSTRAINT FK_Ttb_categorie__Ttb_sezioni "& vbCrLf & _
   	  "		FOREIGN KEY (ca_sezione_id) REFERENCES Ttb_sezioni (se_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 118)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 119
'...........................................................................................
'aggiunge campi per NEXTtravel
'...........................................................................................
sql = "ALTER TABLE Ttb_categorie ADD COLUMN "& vbCrLf & _
	  "		ca_link VARCHAR(255) NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_sezioni ADD COLUMN "& vbCrLf & _
	  "		se_link VARCHAR(255) NULL; "
CALL DB.Execute(sql, 119)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 120
'...........................................................................................
'aggiunge campi per NEXTtravel
'...........................................................................................
sql = "ALTER TABLE Ttb_sezioni ADD COLUMN "& vbCrLf & _
	  "		se_nextWeb_ps INTEGER NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_categorie ADD COLUMN "& vbCrLf & _
	  "		ca_nextWeb_ps INTEGER NULL; "& vbCrLf & _
	  "ALTER TABLE Ttb_sottoCategorie ADD COLUMN "& vbCrLf & _
	  "		sc_nextWeb_ps INTEGER NULL; "
CALL DB.Execute(sql, 120)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 121
'...........................................................................................
'aggiunge campi per NEXTtravel
'...........................................................................................
sql = "ALTER TABLE Ttb_sottoCategorie ADD COLUMN "& vbCrLf & _
	  "		sc_default BIT NULL; "& vbCrLf & _
	  "INSERT INTO Ttb_sottoCategorie(sc_nome_it, sc_nome_en, sc_nome_fr, sc_nome_es, sc_nome_de, sc_categoria_id, sc_default) "& vbCrLf & _
	  "SELECT ca_nome_it, ca_nome_en, ca_nome_fr, ca_nome_es, ca_nome_de, ca_id, 1 FROM Ttb_categorie"
CALL DB.Execute(sql, 121)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 122
'...........................................................................................
'aggiunge campi per NEXTtravel
'...........................................................................................
sql = "ALTER TABLE Ttb_viaggi ADD COLUMN "& vbCrLf & _
	  "		vi_partenza_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_partenza_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_partenza_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_partenza_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_partenza_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		vi_notti_it VARCHAR(100) NULL, "& vbCrLf & _
	  "		vi_notti_en VARCHAR(100) NULL, "& vbCrLf & _
	  "		vi_notti_fr VARCHAR(100) NULL, "& vbCrLf & _
	  "		vi_notti_es VARCHAR(100) NULL, "& vbCrLf & _
	  "		vi_notti_de VARCHAR(100) NULL; "& vbCrLf & _
	  "UPDATE Ttb_viaggi SET vi_partenza_it = vi_partenza, vi_notti_it = vi_notti;"& vbCrLf & _
	  "ALTER TABLE Ttb_viaggi DROP COLUMN vi_partenza, vi_notti"
CALL DB.Execute(sql, 122)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 123
'...........................................................................................
'aggiunggiornamento per variazione struttura directory frameworks
'...........................................................................................
sql = "UPDATE tb_siti SET sito_dir='NEXTpassport' WHERE id_sito=1 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTweb' WHERE id_sito=2 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTcom' WHERE id_sito=3 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTnews' WHERE id_sito=4 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTlink' WHERE id_sito=5 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTmenu' WHERE id_sito=6 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTflat' WHERE id_sito=7 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTmemo' WHERE id_sito=8 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTbanner' WHERE id_sito=9 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTclub' WHERE id_sito=10 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTbooking' WHERE id_sito=11 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTguestbook' WHERE id_sito=12 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTcontract' WHERE id_sito=13 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTfaq' WHERE id_sito=14 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTteam' WHERE id_sito=15 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTbookingportal' WHERE id_sito=16 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTflatportal' WHERE id_sito=17 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTrealestate' WHERE id_sito=18 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTb2b' WHERE id_sito=19 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTschool' WHERE id_sito=20 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTcongress' WHERE id_sito=21 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTb2b_import' WHERE id_sito=22 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTb2b_mailing' WHERE id_sito=23 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTtravel' WHERE id_sito=24 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='NEXTweb4' WHERE id_sito=25 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../prodotti_PANIZZI' WHERE id_sito=100 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='APTadmin' WHERE id_sito=101 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../foto/admin' WHERE id_sito=102 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='APTaontrolloq' WHERE id_sito=103 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../APTbussola' WHERE id_sito=104 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../leo/admin' WHERE id_sito=105 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../prenotazioni' WHERE id_sito=106 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../venicecard/admin' WHERE id_sito=107 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../prenotazioniVENEZIASI' WHERE id_sito=108 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../widmann' WHERE id_sito=109 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../prenotazioni_TURIVE' WHERE id_sito=110 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../prenotazioni_UNI' WHERE id_sito=111 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../mag_centrale' WHERE id_sito=112 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../mag_omaggistica' WHERE id_sito=113 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../mag_spedizioni' WHERE id_sito=114 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../presenze' WHERE id_sito=115 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../APTdistribuzione' WHERE id_sito=116 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../bookings' WHERE id_sito=117 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../datiturismo' WHERE id_sito=118 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../agenzie' WHERE id_sito=119 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='timetable' WHERE id_sito=120 ;" + vbCrLf + _
	  "UPDATE tb_siti SET sito_dir='../EuropeAssistance' WHERE id_sito=121 ;"
CALL DB.Execute(sql, 123)
'*******************************************************************************************


'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
' AGGIORNAMENTI PER RIMOZIONE VECCHIA STRUTTURA B2B
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 124
'...........................................................................................
'rimozione query non utilizzata: ATTENZIONE NON IN TUTTI I DATABASE E' PRESENTE, POTREBBE
'DARE ERRORE L'ESECUZIONE!!
'...........................................................................................
sql = "DROP VIEW qry_email_per_dipendente"
CALL DB.ProtectedExecute(sql, 124, true)
'*******************************************************************************************


'*******************************************************************************************
'rimozione tabelle non utilizzate nel vecchio b2b

'...........................................................................................
'AGGIORNAMENTO 125
'...........................................................................................
sql = "DROP TABLE tb_logs; " + _
	  "DROP TABLE tb_Categorie; " + _
	  "DROP TABLE tb_uffici; " + _
	  "DROP TABLE tb_amministratori; "
CALL DB.Execute(sql, 125)

sql = " ALTER TABLE tb_valori DROP CONSTRAINT tb_variantitb_valori; " + _
	  " DROP TABLE tb_varianti; "
CALL DB.Execute(sql, 126)

sql = " ALTER TABLE rel_cart_var DROP CONSTRAINT tb_valorirel_cart_var; " + _
	  " ALTER TABLE rel_cart_var DROP CONSTRAINT tb_dett_cartrel_cart_var; " + _
	  " DROP TABLE rel_cart_var; "
CALL DB.Execute(sql, 127)

sql = " ALTER TABLE rel_det_var DROP CONSTRAINT tb_valorirel_det_var; " + _
	  " ALTER TABLE rel_det_var DROP CONSTRAINT tb_dettagli_ordrel_dett_var; " + _
	  " DROP TABLE rel_det_var; "
CALL DB.Execute(sql, 128)

sql = " ALTER TABLE rel_art_valori DROP CONSTRAINT tb_valorirel_art_valori; " + _
	  " ALTER TABLE rel_art_valori DROP CONSTRAINT tb_articolirel_art_valori; " + _
	  " DROP TABLE rel_art_valori ;"
CALL DB.Execute(sql, 129)

sql = " ALTER TABLE rel_art_foto DROP CONSTRAINT tb_articolirel_art_foto; " + _
	  " DROP TABLE rel_art_foto; "
CALL DB.Execute(sql, 130)

sql = " ALTER TABLE rel_art_acc DROP CONSTRAINT [{8AF27A47-BABB-43E6-8CE1-E602D042A81D}]; " + _
	  " DROP TABLE rel_art_acc; "
CALL DB.Execute(sql, 131)

sql = " ALTER TABLE rel_art_ctech DROP CONSTRAINT [{6257B7F1-610F-4762-A403-F6E6CD04E996}]; " + _
	  " ALTER TABLE rel_art_ctech DROP CONSTRAINT tb_articolirel_art_ctech; " + _
	  " DROP TABLE rel_art_ctech; "
CALL DB.Execute(sql, 132)

sql = " ALTER TABLE tb_dettagli_ord DROP CONSTRAINT tb_articolitb_dettagli_prev; " + _
	  " ALTER TABLE tb_dettagli_ord DROP CONSTRAINT [{25A8D1C0-72A3-4A34-B7F4-8FEDAA7B7E51}]; " + _
	  " DROP TABLE tb_dettagli_ord; "
CALL DB.Execute(sql, 133)

sql = " ALTER TABLE tb_codici DROP CONSTRAINT [{5C36B9D3-4988-4D95-9DC2-C1A716B23CCB}]; " + _
	  " ALTER TABLE tb_codici DROP CONSTRAINT tb_articolitb_codici; " + _
	  " DROP TABLE tb_codici; "
CALL DB.Execute(sql, 134)

sql = " ALTER TABLE tb_elenco_prezzi DROP CONSTRAINT [{DDEC6E2C-720F-4F37-83B1-6E0D77588F5A}]; " + _
	  " ALTER TABLE tb_elenco_prezzi DROP CONSTRAINT tb_articolitb_elenco_prezzi; " + _
	  " DROP TABLE tb_elenco_prezzi; "
CALL DB.Execute(sql, 135)

sql = " ALTER TABLE tb_dett_Cart DROP CONSTRAINT [{DE780202-6909-4A75-A922-F42CBDCB20D8}]; " + _
	  " ALTER TABLE tb_Dett_Cart DROP CONSTRAINT tb_articolitb_dett_cart; " + _
	  " DROP TABLE tb_Dett_Cart; "
CALL DB.Execute(sql, 136)

sql = " ALTER TABLE tb_ordini DROP CONSTRAINT [{E26259FC-90F2-41CE-91C0-CCFBC39CB463}]; " + _
	  " DROP TABLE tb_ordini; "
CALL DB.Execute(sql, 137)

sql = " ALTER TABLE tb_Articoli DROP CONSTRAINT tb_lineetb_articoli; " + _
	  " DROP TABLE tb_articoli; "
CALL DB.Execute(sql, 138)

sql = " ALTER TABLE tb_linee DROP CONSTRAINT tb_cat_merctb_linee; " + _
	  " DROP TABLE tb_linee;"
CALL DB.Execute(sql, 139)

sql = " ALTER TABLE tb_cat_merc DROP CONSTRAINT tb_aziendetb_cat_merc; " + _
	  " DROP TABLE tb_cat_merc; "
CALL DB.Execute(sql, 140)

sql = " ALTER TABLE tb_shopping_cart DROP CONSTRAINT tb_rivenditoritb_shopping_cart; " + _
	  " DROP TABLE tb_shopping_cart; "
CALL DB.Execute(sql, 141)

sql = " ALTER TABLE tb_carattech DROP CONSTRAINT tb_aziendetb_carattech; " + _
	  " DROP TABLE tb_carattech; "
CALL DB.Execute(sql, 142)

sql = " ALTER TABLE tb_rivenditori DROP CONSTRAINT tb_agentitb_rivenditori; " + _
	  " ALTER TABLE tb_rivenditori DROP CONSTRAINT tb_aziendetb_rivenditori; " + _
	  " ALTER TABLE tb_rivenditori DROP CONSTRAINT [{2BB3DE3F-1BAC-42AC-898D-71D186D38433}]; " + _
	  " ALTER TABLE tb_rivenditori DROP CONSTRAINT tb_listinitb_rivenditori; " + _
	  " ALTER TABLE tb_rivenditori DROP CONSTRAINT tb_monetetb_rivenditori; " + _
	  " DROP TABLE tb_rivenditori; "
CALL DB.Execute(sql, 143)

sql = " ALTER TABLE tb_agenti DROP CONSTRAINT tb_aziendetb_agenti; " + _
	  " DROP TABLE tb_agenti; "
CALL DB.Execute(sql, 144)

sql = " DROP TABLE tb_valute; " + _
	  " DROP TABLE tb_listini; " + _
	  " DROP TABLE tb_lista_codici; " + _
	  " DROP TABLE tb_valori;" + _
	  " DROP TABLE tb_aziende; "
CALL DB.Execute(sql, 145)
'*******************************************************************************************

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 146
'...........................................................................................
'corregge nome applicatiivi nextCom e nextDoc+
'...........................................................................................
sql = "UPDATE tb_siti SET sito_nome='" + IIF(Application("NextCrm"), "NEXT-doc+ [comunicazioni &amp;; documenti]", "NEXT-com [gestione comunicazioni]") + "' WHERE id_sito=3; "
CALL DB.Execute(sql, 146)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 147
'...........................................................................................
'aggiunge campo alla tabella tipi camere next-Booking
'...........................................................................................
sql = "ALTER TABLE btb_tipiCamera ADD tipC_numero int NULL;"
CALL DB.Execute(sql, 147)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 148
'...........................................................................................
'corregge nome applicativo next-web 4
'...........................................................................................
sql = "UPDATE tb_siti SET sito_nome='NEXT-web 4.0 [gestione grafica e contenuti]', sito_dir='NEXTweb4' WHERE id_sito=25"
CALL DB.Execute(sql, 148)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 149
'...........................................................................................
'aggiunge campi per NEXTtravel: nuovi campi su richiesta informazioni
'...........................................................................................
sql = " ALTER TABLE Trel_info ADD COLUMN " + _
	  "		in_citta_partenza TEXT(250) WITH COMPRESSION NULL, " + _
	  " 	in_periodo_dal DATETIME NULL, " + _
  	  " 	in_periodo_al DATETIME NULL, " + _
	  "		in_partecipanti INT NULL "
CALL DB.Execute(sql, 149)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 150
'...........................................................................................
'aggiunge campi per NEXTtravel: nuovi campi su richiesta informazioni
'...........................................................................................
sql = " ALTER TABLE Trel_info DROP COLUMN in_periodo_dal; " + _
	  " ALTER TABLE Trel_info DROP COLUMN in_periodo_al; " + _
	  " ALTER TABLE Trel_info ADD COLUMN in_periodo TEXT(250) WITH COMPRESSION NULL "
CALL DB.Execute(sql, 150)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 151
'...........................................................................................
' Aggiornamento delle procedura nextmemo per passaggio alla versione con categorie
'...........................................................................................
sql = "CREATE TABLE tb_categorieCircolari ("& vbCrLf & _
	  "		catC_id COUNTER CONSTRAINT PK_Ttb_categorieCircolari PRIMARY KEY, "& vbCrLf & _
	  "		catC_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		catC_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		catC_codice VARCHAR(50) NULL, "& vbCrLf & _
	  "		catC_foglia BIT NULL, "& vbCrLf & _
	  "		catC_livello INTEGER NULL, "& vbCrLf & _
	  "		catC_padre_id INTEGER NULL, "& vbCrLf & _
	  "		catC_ordine INTEGER NULL, "& vbCrLf & _
	  "		catC_ordine_assoluto INTEGER NULL, "& vbCrLf & _
	  "		catC_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		catC_descr_en TEXT WITH COMPRESSION NULL "& vbCrLf & _
	  ");"& vbCrLf &_
 	  " ALTER TABLE tb_Circolari ADD COLUMN "& vbCrLf & _
	  "		CI_idcategoria INT NULL "& vbCrLf
CALL DB.Execute(sql, 151)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 152
'...........................................................................................
'aggiunge campi per NEXTtravel: nuovi campi su categorie e sottocategorie per visibile / non visibile
'...........................................................................................
sql = " ALTER TABLE Ttb_categorie ADD COLUMN ca_visibile BIT NULL ;" + _
	  " UPDATE Ttb_categorie SET ca_visibile=1 ;" + _
	  " ALTER TABLE Ttb_SottoCategorie ADD COLUMN sc_visibile BIT NULL; " + _
	  " UPDATE Ttb_SottoCategorie SET sc_visibile=1 "
CALL DB.Execute(sql, 152)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 153
'...........................................................................................
'aggiunge campi per NEXTlink: ordine
'...........................................................................................
sql = " ALTER TABLE tb_links ADD COLUMN link_ordine INTEGER NULL;" + _
      " UPDATE tb_links SET link_ordine = 0"
CALL DB.Execute(sql, 153)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 154
'...........................................................................................
'aggiunge campi per descrittori: ordine, flag principale
'...........................................................................................
sql = " ALTER TABLE tb_descrittori ADD COLUMN " + _
	  "		descr_ordine INTEGER NULL, " + _
	  " 	descr_principale BIT NULL; " + _
	  " UPDATE tb_descrittori SET descr_ordine = 0 "
CALL DB.Execute(sql, 154)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 155
'...........................................................................................
'aggiunge tabelle e relazioni per CensimentoImmobiliare
'...........................................................................................
sql = "CREATE TABLE CItb_immobili ("& vbCrLf & _
	  "		im_ID COUNTER CONSTRAINT PK_CItb_immobili PRIMARY KEY, "& vbCrLf & _
	  "		im_nome_it VARCHAR(150) NULL, "& vbCrLf & _
	  "		im_nome_en VARCHAR(150) NULL, "& vbCrLf & _
	  "		im_nome_fr VARCHAR(150) NULL, "& vbCrLf & _
	  "		im_nome_es VARCHAR(150) NULL, "& vbCrLf & _
	  "		im_nome_de VARCHAR(150) NULL, "& vbCrLf & _
	  "		im_indirizzo VARCHAR(255) NULL, "& vbCrLf & _
	  "		im_cap INTEGER NULL, "& vbCrLf & _
	  "		im_citta VARCHAR(50) NULL, "& vbCrLf & _
	  "		im_localita VARCHAR(100) NULL, "& vbCrLf & _
	  "		im_provincia VARCHAR(50) NULL, "& vbCrLf & _
	  "		im_dataIns DATETIME NULL, "& vbCrLf & _
	  "		im_dataUM DATETIME NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE CItb_descrittori ("& vbCrLf & _
	  "		des_ID COUNTER CONSTRAINT PK_CItb_descrittori PRIMARY KEY, "& vbCrLf & _
	  "		des_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_tipo INTEGER NULL, "& vbCrLf & _
	  "		des_ordine INTEGER NULL, "& vbCrLf & _
	  "		des_principale BIT NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE CIrel_descrittori_immobili ("& vbCrLf & _
	  "		rdi_ID COUNTER CONSTRAINT PK_CIrel_descrittori_immobili PRIMARY KEY, "& vbCrLf & _
	  "		rdi_descrittore_id INTEGER NULL, "& vbCrLf & _
	  "		rdi_immobile_id INTEGER NULL, "& vbCrLf & _
	  "		rdi_valore_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_de VARCHAR(255) NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE CIrel_descrittori_immobili ADD CONSTRAINT FK_CIrel_descrittori_immobili__CItb_descrittori "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_descrittore_id) REFERENCES CItb_descrittori (des_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE CIrel_descrittori_immobili ADD CONSTRAINT FK_CIrel_descrittori_viaggi__CItb_viaggi "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_immobile_id) REFERENCES CItb_immobili (im_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 155)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 156
'...........................................................................................
'Censimento immobili: modifica campo cap
'...........................................................................................
sql = " ALTER TABLE CItb_immobili ALTER COLUMN " + _
	  "		im_cap VARCHAR(20) NULL"
CALL DB.Execute(sql, 156)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 157
'...........................................................................................
'news: aggiunge campo url esterno.
'...........................................................................................
sql = " ALTER TABLE tb_news ADD " + _
	  " 	news_url TEXT(250) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 157)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 158
'...........................................................................................
'aggiorna campi parametri applicativi
'...........................................................................................
sql = " ALTER TABLE tb_siti_parametri ALTER COLUMN " + _
	  "		par_key TEXT(250) WITH COMPRESSION NOT NULL; " + _
	  " ALTER TABLE tb_siti_parametri ALTER COLUMN " + _
	  "		par_value TEXT(250) WITH COMPRESSION NOT NULL "
CALL DB.Execute(sql, 158)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 159
'...........................................................................................
'aggiunge campi per uniformare categorie NEXTmemo a nuova classe
'...........................................................................................
sql = " ALTER TABLE tb_categorieCircolari ADD "+ _
	  "		catC_tipologia_padre_base INT NULL, "+ _
	  "		catC_visibile BIT NULL, "+ _
	  "		catC_albero_visibile BIT NULL; "+ _
	  " ALTER TABLE tb_categorieCircolari ALTER COLUMN "+ _
	  " 	catC_ordine_assoluto TEXT(250) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 159)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 160
'...........................................................................................
'aggiorna dati visibilita' e visibilita' ramo dell'albero NEXT-memo
'...........................................................................................
sql = " UPDATE tb_categorieCircolari SET catC_visibile=1, catC_albero_visibile=1 "
CALL DB.Execute(sql, 160)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 161
'...........................................................................................
'aggiorna NEXTbanner
'...........................................................................................
sql = " ALTER TABLE tb_banner ADD "+ _
	  " 	ban_param TEXT(100) WITH COMPRESSION NULL, "+ _
	  "		ban_value TEXT(250) WITH COMPRESSION NULL; "+ _
	  " ALTER TABLE tb_tipiBanner ADD "+ _
	  "		tipoB_note TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 161)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 162
'...........................................................................................
'aggiorna NEXT-tr4avel per aggiunta campi descrittori
'...........................................................................................
sql = " ALTER TABLE Trel_descrittori_viaggi ADD "+ _
	  "		rdv_memo_it TEXT WITH COMPRESSION NULL, " + _
	  "		rdv_memo_en TEXT WITH COMPRESSION NULL, " + _
	  "		rdv_memo_fr TEXT WITH COMPRESSION NULL, " + _
	  "		rdv_memo_de TEXT WITH COMPRESSION NULL, " + _
	  "		rdv_memo_es TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 162)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 163
'...........................................................................................
'aggiorna NEXT-travel per spostamento dati su nuovi campi
'...........................................................................................
sql = " UPDATE Trel_descrittori_viaggi SET " + _
	  "		rdv_memo_it=rdv_valore_it, " + _
	  "		rdv_memo_en=rdv_valore_en, " + _
	  "		rdv_memo_fr=rdv_valore_fr, " + _
	  "		rdv_memo_de=rdv_valore_de, " + _
	  "		rdv_memo_es=rdv_valore_es " + _
	  " WHERE rdv_descrittore_id IN (SELECT des_id FROM Ttb_descrittori WHERE des_tipo=201) "
CALL DB.Execute(sql, 163)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 164
'...........................................................................................
'crea NEXT-gallery per la gestione di un set di immagini suddiviso per categorie
'...........................................................................................
sql = "CREATE TABLE ptb_gallery ( " &_
		"gallery_id COUNTER CONSTRAINT PK_gallery_id PRIMARY KEY, " &_
		"gallery_name_it varchar(250) NULL, " &_
		"gallery_name_en varchar(250) NULL, " &_
		"gallery_name_fr varchar(250) NULL, " &_
		"gallery_name_de varchar(250) NULL, " &_
		"gallery_name_es varchar(250) NULL, " &_
		"gallery_no INTEGER, " &_
		"config_enabled BIT NULL " &_
		"); " & vbCRLF & _
		"CREATE TABLE ptb_categorieGallery ("& vbCrLf & _
	  "		catC_id COUNTER CONSTRAINT PK_ptb_categorieGallery PRIMARY KEY, "& vbCrLf & _
	  "		catC_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		catC_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		catC_codice VARCHAR(50) NULL, "& vbCrLf & _
	  "		catC_foglia BIT NULL, "& vbCrLf & _
	  "		catC_livello INTEGER NULL, "& vbCrLf & _
	  "		catC_padre_id INTEGER NULL, "& vbCrLf & _
	  "		catC_ordine INTEGER NULL, "& vbCrLf & _
	  "		catC_ordine_assoluto INTEGER NULL, "& vbCrLf & _
	  "		catC_descr_it TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	  "		catC_descr_en TEXT WITH COMPRESSION NULL "& vbCrLf & _
	  ");"& vbCrLf &_
 	  " ALTER TABLE ptb_gallery ADD COLUMN "& vbCrLf & _
	  "		gallery_idcategoria INT NULL; "& vbCrLf & _
	  "ALTER TABLE ptb_gallery ADD CONSTRAINT FK_ptb_gallery__ptb_categorieGallery "& vbCrLf & _
   	  "		FOREIGN KEY (gallery_idcategoria) REFERENCES ptb_categorieGallery (catC_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "	CREATE TABLE ptb_descrittori ("& vbCrLf & _
	  "		des_ID COUNTER CONSTRAINT PK_ptb_descrittori PRIMARY KEY, "& vbCrLf & _
	  "		des_nome_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	  "		des_tipo INTEGER NULL, "& vbCrLf & _
	  "		des_ordine INTEGER NULL, "& vbCrLf & _
	  "		des_principale BIT NULL"& vbCrLf & _
	  ");"& vbCrLf & _
	  "CREATE TABLE prel_descrittori_gallery ("& vbCrLf & _
	  "		rdi_ID COUNTER CONSTRAINT PK_prel_descrittori_gallery PRIMARY KEY, "& vbCrLf & _
	  "		rdi_descrittore_id INTEGER NULL, "& vbCrLf & _
	  "		rdi_gallery_id INTEGER NULL, "& vbCrLf & _
	  "		rdi_valore_it VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_en VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_fr VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_es VARCHAR(255) NULL, "& vbCrLf & _
	  "		rdi_valore_de VARCHAR(255) NULL "& vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE prel_descrittori_gallery ADD CONSTRAINT FK_prel_descrittori_gallery__ptb_descrittori "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_descrittore_id) REFERENCES ptb_descrittori (des_ID) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"& vbCrLf & _
	  "ALTER TABLE prel_descrittori_gallery ADD CONSTRAINT FK_prel_descrittori_gallery__ptb_gallery "& vbCrLf & _
   	  "		FOREIGN KEY (rdi_gallery_id) REFERENCES ptb_gallery (gallery_id) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;" & vbCrLf & _
	  "CREATE TABLE ptb_Immagini( " & _
	  "	I_Id COUNTER CONSTRAINT PK_ptb_Immagini PRIMARY KEY, " & _
	  "	I_Gallery_id INT NULL, "& vbCrLf & _
	  " I_Titolo TEXT(250) WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Didascalia_it TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Didascalia_en TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Didascalia_fr TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Didascalia_de TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Didascalia_es TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Pubblicazione DATETIME NULL, " & vbCrLf & _
	  " I_File TEXT WITH COMPRESSION NULL, " & vbCrLf & _
	  " I_Visibile BIT" & vbCrLf & _
	  ");"& vbCrLf & _
	  "ALTER TABLE ptb_Immagini ADD CONSTRAINT FK_ptb_Immagini__ptb_gallery "& vbCrLf & _
   	  "		FOREIGN KEY (I_Gallery_id) REFERENCES ptb_gallery (gallery_ID) "& vbCrLf & _
	  " 	ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 164)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 165
'...........................................................................................
'aggiunge campi per uniformare categorie NEXTgallery a nuova classe
'...........................................................................................
sql = " ALTER TABLE ptb_categorieGallery ADD "+ _
	  "		catC_tipologia_padre_base INT NULL, "+ _
	  "		catC_visibile BIT NULL, "+ _
	  "		catC_albero_visibile BIT NULL; "+ _
	  " ALTER TABLE ptb_categorieGallery ALTER COLUMN "+ _
	  " 	catC_ordine_assoluto TEXT(250) WITH COMPRESSION NULL"
CALL DB.Execute(sql, 165)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 166
'...........................................................................................
'aggiunge campi per altre lingue a NEXTgallery
'...........................................................................................
sql = " ALTER TABLE ptb_categorieGallery ADD "+ _
	 "		catC_nome_fr VARCHAR(255) NULL, "& vbCrLf & _
	 "		catC_nome_de VARCHAR(255) NULL, "& vbCrLf & _
	 "		catC_nome_es VARCHAR(255) NULL, "& vbCrLf & _
	 "		catC_descr_fr TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	 "		catC_descr_de TEXT WITH COMPRESSION NULL, "& vbCrLf & _
	 "		catC_descr_es TEXT WITH COMPRESSION NULL; "
CALL DB.Execute(sql, 166)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 167
'...........................................................................................
' rimuovi config_enabled
'...........................................................................................
sql = " ALTER TABLE ptb_gallery DROP COLUMN  "+ _
	 "		config_enabled; "
CALL DB.Execute(sql, 167)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 168
'...........................................................................................
' rimuovi config_enabled
'...........................................................................................
sql = " ALTER TABLE ptb_gallery ADD "+ _
	"gallery_visibile BIT NULL, " & vbCrLf &_
	"gallery_codice VARCHAR(50) NULL, "& vbCrLf & _
	"gallery_ordine INT NULL; "
CALL DB.Execute(sql, 168)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 169
'...........................................................................................
'aggiorna NEXT-gallery per aggiunta campi descrittori
'...........................................................................................
sql = " ALTER TABLE prel_descrittori_gallery ADD "+ _
	  "		rdi_memo_it TEXT WITH COMPRESSION NULL, " + _
	  "		rdi_memo_en TEXT WITH COMPRESSION NULL, " + _
	  "		rdi_memo_fr TEXT WITH COMPRESSION NULL, " + _
	  "		rdi_memo_de TEXT WITH COMPRESSION NULL, " + _
	  "		rdi_memo_es TEXT WITH COMPRESSION NULL "
CALL DB.Execute(sql, 169)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 170
'...........................................................................................
'aggiorna NEXT-gallery per aggiunta campi tabella immagini
'...........................................................................................
sql = " ALTER TABLE ptb_Immagini DROP COLUMN I_File; " + _
	  " ALTER TABLE ptb_Immagini ADD "+ _
	  "		I_numero INT NULL, " & vbCrLf &_
	  "		I_thumb VARCHAR(255) NULL, "& vbCrLf & _
	  "		I_zoom VARCHAR(255) NULL, "& vbCrLf & _
	  "		I_ordine INT NULL; "
CALL DB.Execute(sql, 170)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 171
'...........................................................................................
'controlla e ripulisce directory vecchie e residue
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 171)
if DB.last_update_executed then
	CALL Aggiornamento__FRAMEWORK_CORE__pulizia_directory(conn, rs)
end if
'...........................................................................................
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 172
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__1(conn)
CALL DB.Execute(sql, 172)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 173
'...........................................................................................
sql = Aggiornamento__BOOKING__1(conn)
CALL DB.Execute(sql, 173)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 174
'...........................................................................................
sql = Aggiornamento__BOOKING__2(conn)
CALL DB.Execute(sql, 174)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 175
'...........................................................................................
sql = Aggiornamento__BOOKING__3(conn)
CALL DB.Execute(sql, 175)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 176
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__2(conn)
CALL DB.Execute(sql, 176)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 177
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__3(conn)
CALL DB.Execute(sql, 177)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 178
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__4(conn)
CALL DB.Execute(sql, 178)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 179
'...........................................................................................
'rimuove struttura dati per censimento immobili. (spostato su database sql prisma)
'...........................................................................................
sql = " ALTER TABLE CIrel_descrittori_immobili DROP CONSTRAINT FK_CIrel_descrittori_immobili__CItb_descrittori; " + _
	  " ALTER TABLE CIrel_descrittori_immobili DROP CONSTRAINT FK_CIrel_descrittori_viaggi__CItb_viaggi; " + _
	  " DROP TABLE CItb_immobili; " + _
	  " DROP TABLE CItb_descrittori; " + _
	  " DROP TABLE CIrel_descrittori_immobili; "
CALL DB.Execute(sql, 179)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 180
'...........................................................................................
sql = rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
CALL DB.Execute(sql, 180)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 181
'...........................................................................................
CALL rebuild__FRAMEWORK_CORE__cartelle(conn, rs, DB, 181)
'*******************************************************************************************
'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 182
'...........................................................................................
sql = Install__BOOKING2(conn)
CALL DB.Execute(sql, 182)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 183
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__5(conn)
CALL DB.Execute(sql, 183)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 184
'...........................................................................................
sql = Aggiornamento__BOOKING2__1(conn)
CALL DB.Execute(sql, 184)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 185
'...........................................................................................
sql = Aggiornamento__BOOKING2__2(conn)
CALL DB.Execute(sql, 185)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 186
'...........................................................................................
sql = Aggiornamento__BOOKING2__3(conn)
CALL DB.Execute(sql, 186)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 187
'...........................................................................................
sql = Aggiornamento__BOOKING2__4(conn)
CALL DB.Execute(sql, 187)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 188
'...........................................................................................
sql = Aggiornamento__FLAT__1(conn)
CALL DB.Execute(sql, 188)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 189
'...........................................................................................
sql = Aggiornamento__BOOKING2__5(conn)
CALL DB.Execute(sql, 189)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 190
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__6(conn)
CALL DB.Execute(sql, 190)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 191
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__7(conn)
CALL DB.Execute(sql, 191)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 192
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__8(conn)
CALL DB.Execute(sql, 192)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 193
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__9(conn)
CALL DB.Execute(sql, 193)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 194
'...........................................................................................
sql = Aggiornamento__BOOKING2__6(conn)
CALL DB.Execute(sql, 194)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 195
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__10(conn)
CALL DB.Execute(sql, 195)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 196
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__11(conn)
CALL DB.Execute(sql, 196)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 197
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__12(DB, rs, 197)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 198
'...........................................................................................
sql = Aggiornamento__BOOKING2__7(conn)
CALL DB.Execute(sql, 198)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 199
'...........................................................................................
sql = Aggiornamento__BOOKING2__8(conn)
CALL DB.Execute(sql, 199)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 200
'...........................................................................................
sql = Aggiornamento__FLAT__2(conn)
CALL DB.Execute(sql, 200)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 201
'...........................................................................................
sql = Aggiornamento__FLAT__3(conn)
CALL DB.Execute(sql, 201)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 202
'...........................................................................................
sql = Aggiornamento__BOOKING2__9(conn)
CALL DB.Execute(sql, 202)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 203
'...........................................................................................
sql = Aggiornamento__BOOKING2__10(conn)
CALL DB.Execute(sql, 203)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 204
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__13(conn)
CALL DB.Execute(sql, 204)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 205
'...........................................................................................
sql = Aggiornamento__BOOKING2__11(conn)
CALL DB.Execute(sql, 205)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 206
'...........................................................................................
sql = Install__BOOKING3(conn) + _
	  Aggiornamento__BOOKING3__1(conn) + _
	  Aggiornamento__BOOKING3__2(conn) + _
	  Aggiornamento__BOOKING3__3(conn) + _
	  Aggiornamento__BOOKING3__4(conn) + _
	  Aggiornamento__BOOKING3__5(conn) + _
	  Aggiornamento__BOOKING3__6(conn) + _
	  Aggiornamento__BOOKING3__7(conn) + _
	  Aggiornamento__BOOKING3__8(conn) + _
	  Aggiornamento__BOOKING3__9(conn) + _
	  Aggiornamento__BOOKING3__10(conn)
CALL DB.Execute(sql, 206)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 207
'...........................................................................................
sql = Aggiornamento__BOOKING3__11(conn)
CALL DB.Execute(sql, 207)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 208
'...........................................................................................
sql = Aggiornamento__BOOKING3__12(conn)
CALL DB.Execute(sql, 208)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 209
'...........................................................................................
sql = Aggiornamento__BOOKING3__13(conn)
CALL DB.Execute(sql, 209)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 210
'...........................................................................................
sql = Aggiornamento__BOOKING3__14(conn)
CALL DB.Execute(sql, 210)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 211
'...........................................................................................
sql = Aggiornamento__BOOKING3__15(conn)
CALL DB.Execute(sql, 211)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 212
'...........................................................................................
sql = Aggiornamento__BOOKING3__16(conn)
CALL DB.Execute(sql, 212)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 213
'...........................................................................................
sql = Aggiornamento__BOOKING3__17(conn)
CALL DB.Execute(sql, 213)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 214
'...........................................................................................
sql = Aggiornamento__BOOKING3__18(conn)
CALL DB.Execute(sql, 214)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 215
'...........................................................................................
sql = Aggiornamento__BOOKING3__19(conn)
CALL DB.Execute(sql, 215)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 216
'...........................................................................................
sql = rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
CALL DB.Execute(sql, 216)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 217
'...........................................................................................
sql = Install__FRAMEWORK_CORE__NEXTWEB5(conn)
CALL DB.Execute(sql, 217)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 218
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__14(conn)
CALL DB.Execute(sql, 218)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 219
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__15(conn)
CALL DB.Execute(sql, 219)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 220
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__16(conn)
CALL DB.Execute(sql, 220)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 221
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__17(conn)
CALL DB.Execute(sql, 221)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 222
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__18(conn)
CALL DB.Execute(sql, 222)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 223
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__19(conn)
CALL DB.Execute(sql, 223)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 224
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__20(conn)
CALL DB.Execute(sql, 224)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 225
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__21(conn)
CALL DB.Execute(sql, 225)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 226
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__22(conn)
CALL DB.Execute(sql, 226)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 227
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__23(conn)
CALL DB.Execute(sql, 227)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 228
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__24(conn)
CALL DB.Execute(sql, 228)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 229
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__25(conn)
CALL DB.Execute(sql, 229)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 230
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__26(conn)
CALL DB.Execute(sql, 230)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 231
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__27(conn)
CALL DB.Execute(sql, 231)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 232
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__28(conn)
CALL DB.Execute(sql, 232)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 233
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__29(conn)
CALL DB.Execute(sql, 233)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 234
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__30(conn)
CALL DB.Execute(sql, 234)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 235
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__31(conn)
CALL DB.Execute(sql, 235)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 236
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__32(conn)
CALL DB.Execute(sql, 236)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 237
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__33(conn)
CALL DB.Execute(sql, 237)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 238
'...........................................................................................
sql = Aggiornamento__BOOKING3__20(conn)
CALL DB.Execute(sql, 238)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 239
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__34(conn)
CALL DB.Execute(sql, 239)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 240
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__35(conn)
CALL DB.Execute(sql, 240)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 241
'...........................................................................................
sql = Aggiornamento__BOOKING3__21(conn)
CALL DB.Execute(sql, 241)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 242
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__36(conn)
CALL DB.Execute(sql, 242)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 243
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__37(conn)
CALL DB.Execute(sql, 243)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 244
'...........................................................................................
sql = AggiornamentoSpeciale__FRAMEWORK_CORE__38(DB, rs, 244)
CALL DB.Execute(sql, 244)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 245
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__39(conn)
CALL DB.Execute(sql, 245)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 246
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__40(conn)
CALL DB.Execute(sql, 246)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 247
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__41(conn)
CALL DB.Execute(sql, 247)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 248
'...........................................................................................
sql = Aggiornamento__BOOKING3__22(conn)
CALL DB.Execute(sql, 248)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 249
'...........................................................................................
sql = Aggiornamento__BOOKING3__23(conn)
CALL DB.Execute(sql, 249)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 250
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__42(conn)
CALL DB.Execute(sql, 250)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 251
'...........................................................................................
sql = Aggiornamento__BOOKING3__24(conn)
CALL DB.Execute(sql, 251)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 252: aggiunto campo soc_non_attivo a ctb_soci (next-club)
'...........................................................................................
sql = "ALTER TABLE ctb_soci ADD " + _
	"		soc_non_attivo BIT NOT NULL;  " + _
	"	UPDATE ctb_soci SET soc_non_attivo = false; "
CALL DB.Execute(sql, 252)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 253
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__43(conn)
CALL DB.Execute(sql, 253)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 254
'...........................................................................................
sql = Aggiornamento__BOOKING3__25(conn)
CALL DB.Execute(sql, 254)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 255
'...........................................................................................
sql = Aggiornamento__BOOKING3__26(conn)
CALL DB.Execute(sql, 255)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 256
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__44(conn)
CALL DB.Execute(sql, 256)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 257
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__45(conn)
CALL DB.Execute(sql, 257)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 258
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__46(conn)
CALL DB.Execute(sql, 258)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 259
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__47(conn)
CALL DB.Execute(sql, 259)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 260
'...........................................................................................
sql = Aggiornamento__BOOKING3__27(conn)
CALL DB.Execute(sql, 260)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 261
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__48(conn)
CALL DB.Execute(sql, 261)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 262
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__49(conn)
CALL DB.Execute(sql, 262)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 263
'...........................................................................................
sql = "SELECT * FROM aa_versione"		'annullo chiamata all'installazione del NEXTCOMMENT
CALL DB.Execute(sql, 263)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 264
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__50(conn)
CALL DB.Execute(sql, 264)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 265
'...........................................................................................
sql = Aggiornamento__BOOKING3__28(conn)
CALL DB.Execute(sql, 265)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 266
'...........................................................................................
sql = Aggiornamento__BOOKING3__29(conn)
CALL DB.Execute(sql, 266)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 267
'...........................................................................................
sql = Aggiornamento__BOOKING3__30(conn)
CALL DB.Execute(sql, 267)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 268
'...........................................................................................
sql = AggiornamentoSpeciale__MEMO__1(DB, rs, 268)
CALL DB.Execute(sql, 268)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 269
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__51(conn)
CALL DB.Execute(sql, 269)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 270
'...........................................................................................
sql = Aggiornamento__BOOKING3__31(conn)
CALL DB.Execute(sql, 270)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 271
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__52(conn)
CALL DB.Execute(sql, 271)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 272
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__53(conn)
CALL DB.Execute(sql, 272)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 273
'...........................................................................................
sql = Aggiornamento__BOOKING3__32(conn)
CALL DB.Execute(sql, 273)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 274
'...........................................................................................
sql = Aggiornamento__REALESTATE__1(conn)
CALL DB.Execute(sql, 274)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 275
'...........................................................................................
sql = Aggiornamento__REALESTATE__2(conn)
CALL DB.Execute(sql, 275)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 276
'...........................................................................................
sql = Aggiornamento__REALESTATE__3(conn)
CALL DB.Execute(sql, 276)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 277
'...........................................................................................
sql = Aggiornamento__REALESTATE__4(conn)
CALL DB.Execute(sql, 277)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 278
'...........................................................................................
sql = Aggiornamento__REALESTATE__5(conn)
CALL DB.Execute(sql, 278)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 279
'...........................................................................................
sql = Aggiornamento__REALESTATE__6(conn)
CALL DB.Execute(sql, 279)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 280
'...........................................................................................
sql = Aggiornamento__REALESTATE__7(conn)
CALL DB.Execute(sql, 280)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 281
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__54(conn)
CALL DB.Execute(sql, 281)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 282
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__55(conn)
CALL DB.Execute(sql, 282)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 283
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__56(conn)
CALL DB.Execute(sql, 283)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 284
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__57(conn)
CALL DB.Execute(sql, 284)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 285
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__58(conn)
CALL DB.Execute(sql, 285)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 286
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__59(conn)
CALL DB.Execute(sql, 286)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 287
'...........................................................................................
sql = Aggiornamento__REALESTATE__8(conn)
CALL DB.Execute(sql, 287)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 288
'...........................................................................................
sql = Aggiornamento__REALESTATE__9(conn)
CALL DB.Execute(sql, 288)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 289
'...........................................................................................
sql = Aggiornamento__REALESTATE__10(conn)
CALL DB.Execute(sql, 289)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 290
'...........................................................................................
sql = Aggiornamento__REALESTATE__11(conn)
CALL DB.Execute(sql, 290)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 291
'...........................................................................................
sql = Aggiornamento__REALESTATE__12(conn)
CALL DB.Execute(sql, 291)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 292
'...........................................................................................
sql = Aggiornamento__REALESTATE__13(conn)
CALL DB.Execute(sql, 292)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 293
'...........................................................................................
sql = Aggiornamento__REALESTATE__14(conn)
CALL DB.Execute(sql, 293)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 294
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__60(conn)
CALL DB.Execute(sql, 294)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 295
'...........................................................................................
sql = Aggiornamento__REALESTATE__15(conn)
CALL DB.Execute(sql, 295)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 296
'...........................................................................................
sql = Aggiornamento__REALESTATE__16(conn)
CALL DB.Execute(sql, 296)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 297
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__61(conn)
CALL DB.Execute(sql, 297)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 298
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__62(conn)
CALL DB.Execute(sql, 298)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 299
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__63(conn)
CALL DB.Execute(sql, 299)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 300
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__64(conn)
CALL DB.Execute(sql, 300)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 301
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__65(conn)
CALL DB.Execute(sql, 301)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 302
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__66(conn)
CALL DB.Execute(sql, 302)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 303
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__67(conn)
CALL DB.Execute(sql, 303)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 304
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__68(conn)
CALL DB.Execute(sql, 304)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 305
'...........................................................................................
sql = Aggiornamento__REALESTATE__17(conn)
CALL DB.Execute(sql, 305)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 306
'...........................................................................................
sql = Aggiornamento__REALESTATE__18(conn)
CALL DB.Execute(sql, 306)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 307
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__69(conn)
CALL DB.Execute(sql, 307)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 308
'...........................................................................................
sql = Aggiornamento__REALESTATE__19(conn)
CALL DB.Execute(sql, 308)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 309
'...........................................................................................
sql = Aggiornamento__REALESTATE__20(conn)
CALL DB.Execute(sql, 309)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(309)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 310
'...........................................................................................
sql = Aggiornamento__REALESTATE__21(conn)
CALL DB.Execute(sql, 310)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 311
'...........................................................................................
sql = Aggiornamento__REALESTATE__22(conn)
CALL DB.Execute(sql, 311)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 312
'...........................................................................................
sql = Aggiornamento__REALESTATE__23(conn)
CALL DB.Execute(sql, 312)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(312)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 313
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__70(conn)
CALL DB.Execute(sql, 313)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 314
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__71(conn)
CALL DB.Execute(sql, 314)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 315
'...........................................................................................
sql = Aggiornamento__REALESTATE__24(conn)
CALL DB.Execute(sql, 315)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 316
'...........................................................................................
sql = Aggiornamento__REALESTATE__25(conn)
CALL DB.Execute(sql, 316)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(316)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 317
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__72(DB, 317)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 318
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__73(conn)
CALL DB.Execute(sql, 318)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 319
'...........................................................................................
sql = Aggiornamento__REALESTATE__26(conn)
CALL DB.Execute(sql, 319)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 320
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__74(conn)
CALL DB.Execute(sql, 320)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 321
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__75(conn)
CALL DB.Execute(sql, 321)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(321)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 322
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__76(conn)
CALL DB.Execute(sql, 322)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 323
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__77(conn)
CALL DB.Execute(sql, 323)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 324
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__78(conn)
CALL DB.Execute(sql, 324)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 325
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__79(conn)
CALL DB.Execute(sql, 325)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 326
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__80(conn)
CALL DB.Execute(sql, 326)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 327
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__81(conn)
CALL DB.Execute(sql, 327)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 328
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__82(conn)
CALL DB.Execute(sql, 328)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 329
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__83(conn)
CALL DB.Execute(sql, 329)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 330
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__84(conn)
CALL DB.Execute(sql, 330)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 331
'...........................................................................................
'corregge errore nella associazione dei contatti alle rubriche rimuovendo i duplicati
'esegue stesse attvita' dell'agiornamento Aggiornamento__FRAMEWORK_CORE__79
'...........................................................................................
sql = "SELECT * FROM AA_Versione"
CALL DB.Execute(sql, 331)
'...........................................................................................
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__331(conn)
end if
'...........................................................................................
sub AggiornamentoSpeciale__331(conn)
	dim rsp, rsr, d, sql
	set rsp = server.createobject("adodb.recordset")
	set rsr = server.createobject("adodb.recordset")
	Set d = CreateObject("Scripting.Dictionary")
	sql = "SELECT idElencoIndirizzi FROM tb_indirizzario "
	rsp.open sql, conn, adOpenDynamic, adLockOptimistic
	
	while not rsp.eof
		d.removeAll
		sql = "SELECT * FROM rel_rub_ind WHERE id_indirizzo = "& rsp("idElencoIndirizzi")
		rsr.open sql, conn, adOpenDynamic, adLockOptimistic
		while not rsr.eof
			if d.Exists(CIntero(rsr("id_rubrica"))) then
				rsr.delete
			else
				d(CIntero(rsr("id_rubrica"))) = true
			end if
			
			rsr.movenext
		wend
		rsr.close
		
		rsp.movenext
	wend
	
	rsp.close
	set rsp = nothing
	set rsr = nothing
	set d = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 332
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__85(conn)
CALL DB.Execute(sql, 332)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 333
'...........................................................................................
sql = Aggiornamento__REALESTATE__27(conn)
CALL DB.Execute(sql, 333)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 334
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__86(conn)
CALL DB.Execute(sql, 334)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 335
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__87(DB, 335)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 336
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__88(conn)
CALL DB.Execute(sql, 336)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 337
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__89(conn)
CALL DB.Execute(sql, 337)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(337)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 338
'...........................................................................................
'CALL AggiornamentoSpeciale__FRAMEWORK_CORE__90(DB, rs, 338)
sql = "SELECT * FROM AA_Versione"
CALL DB.Execute(sql, 338)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 339
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__91(conn)
CALL DB.Execute(sql, 339)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(339)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 340
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__92(conn)
CALL DB.Execute(sql, 340)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 341
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__93(conn)
CALL DB.Execute(sql, 341)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 342
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__94(conn)
CALL DB.Execute(sql, 342)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 343
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__95(conn)
CALL DB.Execute(sql, 343)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 344
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__96(conn)
CALL DB.Execute(sql, 344)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 345
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__97(conn)
CALL DB.Execute(sql, 345)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 346
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__98(conn)
CALL DB.Execute(sql, 346)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 347
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__99(conn)
CALL DB.Execute(sql, 347)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 348
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__100(conn)
CALL DB.Execute(sql, 348)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(348)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 349
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__101(conn)
CALL DB.Execute(sql, 349)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 350
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__102(conn)
CALL DB.Execute(sql, 350)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 351
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__103(conn)
CALL DB.Execute(sql, 351)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(351)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 352
'...........................................................................................
sql = Aggiornamento__REALESTATE__28(conn)
CALL DB.Execute(sql, 352)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 353
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__104(conn)
CALL DB.Execute(sql, 353)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 354
'...........................................................................................
sql = Aggiornamento__REALESTATE__29(conn)
CALL DB.Execute(sql, 354)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 355
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__105(conn)
CALL DB.Execute(sql, 355)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 356
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__106(conn)
CALL DB.Execute(sql, 356)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 357
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__107(conn)
CALL DB.Execute(sql, 357)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 358
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__108(conn)
CALL DB.Execute(sql, 358)
'*******************************************************************************************


'*******************************************************************************************
'CALL DB.RebuildIndex_RefreshContents("tb_webs", "id_webs")
'*******************************************************************************************

'*******************************************************************************************
'CALL DB.RebuildIndex_RefreshContents("tb_paginesito", "id_pagineSito")
'*******************************************************************************************

'*******************************************************************************************
'CALL DB.RebuildIndex_RefreshContents("tb_contents_index", "idx_livello")
'*******************************************************************************************

'*******************************************************************************************
'CALL DB.RebuildIndex_OperazioniRicorsive()
'*******************************************************************************************

'?????????
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(358)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 359
'...........................................................................................
sql = Aggiornamento__REALESTATE__30(conn)
CALL DB.Execute(sql, 359)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 360
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__109(conn)
CALL DB.Execute(sql, 360)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 361
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__110(conn)
CALL DB.Execute(sql, 361)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 362
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__111(conn)
CALL DB.Execute(sql, 362)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 363
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__112(conn)
CALL DB.Execute(sql, 363)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 364
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__113(conn)
CALL DB.Execute(sql, 364)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 365
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__114(DB, rs, 365)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 366
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 366)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 367
'riaggiorna la vista per dell'aggiornamento 115
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 367)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(367)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 368
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__116(conn) + _
	  Aggiornamento__FRAMEWORK_CORE__116_bis(conn)
CALL DB.Execute(sql, 368)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 369
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__117(conn)
CALL DB.Execute(sql, 369)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 370
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__118(conn)
CALL DB.Execute(sql, 370)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 371
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__119(conn)
CALL DB.Execute(sql, 371)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 372
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__120(conn)
CALL DB.Execute(sql, 372)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 373
'...........................................................................................
sql = Aggiornamento__REALESTATE__31(conn)
CALL DB.Execute(sql, 373)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 374
'...........................................................................................
sql = "SELECT * FROM AA_Versione"
CALL DB.Execute(sql, 374)
'CALL AggiornamentoSpeciale__FRAMEWORK_CORE__121(DB, rs, 374)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 375
'...........................................................................................
sql = Aggiornamento__REALESTATE__32(conn)
CALL DB.Execute(sql, 375)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 376
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__122(conn)
CALL DB.Execute(sql, 376)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 377
'...........................................................................................
sql = Aggiornamento__GUESTBOOK__1(conn)
CALL DB.Execute(sql, 377)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 378
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__123(conn)
CALL DB.Execute(sql, 378)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 379
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__124(conn)
CALL DB.Execute(sql, 379)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 380
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__125(conn)
CALL DB.Execute(sql, 380)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 381
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__126(conn)
CALL DB.Execute(sql, 381)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 382
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__127(conn)
CALL DB.Execute(sql, 382)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 383
'...........................................................................................
sql = Aggiornamento__BOOKING3__33(conn)
CALL DB.Execute(sql, 383)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 384
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__128(conn)
CALL DB.Execute(sql, 384)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 385
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__129(conn)
CALL DB.Execute(sql, 385)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 386
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__130(conn)
CALL DB.Execute(sql, 386)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 387
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__131(conn)
CALL DB.Execute(sql, 387)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 388
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__132(conn)
CALL DB.Execute(sql, 388)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 389
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__133(conn)
CALL DB.Execute(sql, 389)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 390
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__134(conn)
CALL DB.Execute(sql, 390)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 391
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__135(conn)
CALL DB.Execute(sql, 391)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 392
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__136(conn)
CALL DB.Execute(sql, 392)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 393
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__137(conn)
CALL DB.Execute(sql, 393)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(393)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 394
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__138(conn)
CALL DB.Execute(sql, 394)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 395
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__139(conn)
CALL DB.Execute(sql, 395)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 396
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__140(conn)
CALL DB.Execute(sql, 396)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 397
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__141(conn)
CALL DB.Execute(sql, 397)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(397)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 398
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__142(conn)
CALL DB.Execute(sql, 398)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(398)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 399
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__143(conn)
CALL DB.Execute(sql, 399)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(399)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 400
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__144(conn)
CALL DB.Execute(sql, 400)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(400)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 401
'...........................................................................................
sql = Aggiornamento__REALESTATE__33(conn)
CALL DB.Execute(sql, 401)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(401)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 402
'...........................................................................................
sql = Aggiornamento__REALESTATE__34(conn)
CALL DB.Execute(sql, 402)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 403
'...........................................................................................
sql = Aggiornamento__REALESTATE__35(conn)
CALL DB.Execute(sql, 403)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 404
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__145(conn)
CALL DB.Execute(sql, 404)
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO 405
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__REALESTATE__36(conn, "ru")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 405)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 406
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "ru")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 406)
if DB.last_update_executed then
	'CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "ru", "russo", "Русский")
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 407
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__REALESTATE__36(conn, "cn")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 407)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 408
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "cn")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 408)
if DB.last_update_executed then
	'CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "cn", "Cinese", "中文")
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(408)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 409
'...........................................................................................
sql = Aggiornamento__MEMO__1(conn)
CALL DB.Execute(sql, 409)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 410
'...........................................................................................
sql = Aggiornamento__REALESTATE__37(conn)
CALL DB.Execute(sql, 410)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(410)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 411
'...........................................................................................
sql = Aggiornamento__REALESTATE__38(conn)
CALL DB.Execute(sql, 411)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(411)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 412
'...........................................................................................
sql = Aggiornamento__REALESTATE__39(conn)
CALL DB.Execute(sql, 412)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(412)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 413
'...........................................................................................
sql = Aggiornamento__REALESTATE__40(conn)
CALL DB.Execute(sql, 413)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__40(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 414
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__REALESTATE__36(conn, "pt")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 414)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 415
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
'sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "pt")
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 415)
if DB.last_update_executed then
	'CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "pt", "Portoghese", "Português")
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(415)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 416
'...........................................................................................
sql = Aggiornamento__REALESTATE__41(conn)
CALL DB.Execute(sql, 416)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(416)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 417
'...........................................................................................
sql = Aggiornamento__REALESTATE__42(conn)
CALL DB.Execute(sql, 417)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__42(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 418
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__148(conn)
CALL DB.Execute(sql, 418)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 419
'...........................................................................................
sql = Aggiornamento__REALESTATE__43(conn)
CALL DB.Execute(sql, 419)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 420
'...........................................................................................
sql = Aggiornamento__REALESTATE__44(conn)
CALL DB.Execute(sql, 420)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__44(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 421
'...........................................................................................
sql = Aggiornamento__REALESTATE__45(conn)
CALL DB.Execute(sql, 421)
if DB.last_update_executed then
	dim rs_age, rs_indi
	set rs_age = Server.CreateObject("ADODB.Recordset")
	set rs_indi = Server.CreateObject("ADODB.Recordset")
	rs_age.open "SELECT * FROM rtb_agenzie", conn, adOpenKeySet, adLockOptimistic, adCmdText
	while not rs_age.eof
		rs_indi.open "SELECT NomeOrganizzazioneElencoIndirizzi FROM tb_Indirizzario WHERE IDElencoIndirizzi =" & rs_age("age_id") , conn, adOpenKeySet, adLockOptimistic, adCmdText
		rs_age("age_marchio_it") = rs_indi("NomeOrganizzazioneElencoIndirizzi")
		rs_age.Update
		rs_indi.close
		rs_age.moveNext
	wend
	rs_age.close
	set rs_age = nothing
	set rs_indi = nothing
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(421)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 422
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__149(conn)
CALL DB.Execute(sql, 422)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(422)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 423
'...........................................................................................
sql = Aggiornamento__REALESTATE__46(conn)
CALL DB.Execute(sql, 423)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__46(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 424
'...........................................................................................
sql = Aggiornamento__MEMO__2(conn)
CALL DB.Execute(sql, 424)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 425
'...........................................................................................
sql = Aggiornamento__BOOKING3__34(conn)
CALL DB.Execute(sql, 425)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 426
'...........................................................................................
sql = Aggiornamento__BOOKING3__35(conn)
CALL DB.Execute(sql, 426)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 427
'...........................................................................................
sql = Aggiornamento__BOOKING3__36(conn)
CALL DB.Execute(sql, 427)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 428
'...........................................................................................
sql = Aggiornamento__REALESTATE__47(conn)
CALL DB.Execute(sql, 428)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 429
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__150(conn)
CALL DB.Execute(sql, 429)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(429)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 430
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__151(conn)
CALL DB.Execute(sql, 430)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 431
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__152(conn)
CALL DB.Execute(sql, 431)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(431)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 432
'...........................................................................................
sql = "ALTER TABLE rtb_strutture ADD st_indirizzo_mappa " + SQL_CharField(Conn, 255) + " NULL; "
sql = sql + Aggiornamento__REALESTATE__48(conn)
''sql = Aggiornamento__REALESTATE__48(conn)
CALL DB.Execute(sql, 432)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 433
'...........................................................................................
sql = Aggiornamento__REALESTATE__49(conn)
CALL DB.Execute(sql, 433)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 434
'...........................................................................................
sql = Aggiornamento__REALESTATE__50(conn)
CALL DB.Execute(sql, 434)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 435
'...........................................................................................
sql = Aggiornamento__REALESTATE__51(conn)
CALL DB.Execute(sql, 435)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__51(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 436
'...........................................................................................
sql = Aggiornamento__REALESTATE__52(conn)
CALL DB.Execute(sql, 436)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__52(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 437
'...........................................................................................
sql = Aggiornamento__REALESTATE__53(conn)
CALL DB.Execute(sql, 437)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 438
'...........................................................................................
sql = Aggiornamento__REALESTATE__54(conn)
CALL DB.Execute(sql, 438)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(438)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 439
'...........................................................................................
sql = Aggiornamento__REALESTATE__55(conn)
CALL DB.Execute(sql, 439)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 440
'...........................................................................................
sql = Aggiornamento__REALESTATE__56(conn)
CALL DB.Execute(sql, 440)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 441
'...........................................................................................
sql = Aggiornamento__REALESTATE__57(conn)
CALL DB.Execute(sql, 441)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 442
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__153(conn)
CALL DB.Execute(sql, 442)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(442)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 443
'...........................................................................................
sql = Aggiornamento__REALESTATE__58(conn)
CALL DB.Execute(sql, 443)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 444
'...........................................................................................
sql = Aggiornamento__REALESTATE__59(conn)
CALL DB.Execute(sql, 444)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 445
'...........................................................................................
sql = Aggiornamento__REALESTATE__60(conn)
CALL DB.Execute(sql, 445)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 446
'...........................................................................................
sql = Aggiornamento__REALESTATE__61(conn)
CALL DB.Execute(sql, 446)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 447
'...........................................................................................
sql = Aggiornamento__REALESTATE__62(conn)
CALL DB.Execute(sql, 447)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 448
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__154(conn)
CALL DB.Execute(sql, 448)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 449
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__155(conn)
CALL DB.Execute(sql, 449)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 450
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__156(conn)
CALL DB.Execute(sql, 450)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 451
'...........................................................................................
sql = Aggiornamento__REALESTATE__63(conn)
CALL DB.Execute(sql, 451)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(451)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 452
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__157(conn)
CALL DB.Execute(sql, 452)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(452)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 453
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__158(conn)
CALL DB.Execute(sql, 453)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 454
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__159(conn)
CALL DB.Execute(sql, 454)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(454)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 455
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__160(conn)
CALL DB.Execute(sql, 455)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(455)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 456
'...........................................................................................
' sql = "ALTER TABLE [tb_webs] DROP CONSTRAINT [FK_....];"
sql = Aggiornamento__FRAMEWORK_CORE__161(conn)
CALL DB.Execute(sql, 456)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 457
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__162(conn)
CALL DB.Execute(sql, 457)
if DB.last_update_executed OR cIntero(GetValueList(conn,NULL,"SELECT versione FROM AA_versione"))=457 then 'condizione in OR aggiunta per i DB fermi come ultimo aggiornamento al 457 (come operarealestate)
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__162(conn)

	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__90(DB, rs, 458)
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__121(DB, rs, 459)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 460
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__163(conn)
CALL DB.Execute(sql, 460)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 461
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__165(conn)
CALL DB.Execute(sql, 461)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 462
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__166(conn)
CALL DB.Execute(sql, 462)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 463
'...........................................................................................
sql = Install__MEMO2(conn)
CALL DB.Execute(sql, 463)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 464
'...........................................................................................
sql = AggiornamentoSpeciale__MEMO2__1(DB, rs, 464)
CALL DB.Execute(sql, 464)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 465
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__167(conn)
CALL DB.Execute(sql, 465)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 466
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__168(conn)
CALL DB.Execute(sql, 466)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 467
' Sergio 29/09/2010 - aggiunge colonne mancanti a 
' tb_contents_index 
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__164(conn)
CALL DB.Execute(sql, 467)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 468
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__169(conn)
CALL DB.Execute(sql, 468)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(468)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 469
'...........................................................................................
sql = Aggiornamento__REALESTATE__65(conn)
CALL DB.Execute(sql, 469)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 470
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__170(conn)
CALL DB.Execute(sql, 470)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(470)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 471
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__171(conn)
CALL DB.Execute(sql, 471)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(471)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 472
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__172(conn)
CALL DB.Execute(sql, 472)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(472)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 473
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__173(conn)
CALL DB.Execute(sql, 473)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(473)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 474
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__174(conn)
CALL DB.Execute(sql, 474)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(474)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 475
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__175(conn)
CALL DB.Execute(sql, 475)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__175(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 476
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__176(conn)
CALL DB.Execute(sql, 476)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(476)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 477
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__177(conn)
CALL DB.Execute(sql, 477)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(477)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 478
'...........................................................................................
sql = Aggiornamento__REALESTATE__66(conn)
CALL DB.Execute(sql, 478)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__66(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 479
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__178(conn)
CALL DB.Execute(sql, 479)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(479)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 480
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__179(conn)
CALL DB.Execute(sql, 480)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 481
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__180(conn)
CALL DB.Execute(sql, 481)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 482
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__181(conn)
CALL DB.Execute(sql, 482)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 483
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__182(conn)
CALL DB.Execute(sql, 483)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 484
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__183(conn)
CALL DB.Execute(sql, 484)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 485
'...........................................................................................
sql = Aggiornamento__MEMO2__1(conn)
CALL DB.Execute(sql, 485)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 486
'...........................................................................................
sql = Aggiornamento__MEMO2__2(conn)
CALL DB.Execute(sql, 486)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__2(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 487
'...........................................................................................
sql = Aggiornamento__MEMO2__3(conn)
CALL DB.Execute(sql, 487)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 488
'...........................................................................................
sql = Aggiornamento__MEMO2__4(conn)
CALL DB.Execute(sql, 488)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__4(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 489
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__184(conn)
CALL DB.Execute(sql, 489)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 490
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__185(conn)
CALL DB.Execute(sql, 490)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 491
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__186(conn)
CALL DB.Execute(sql, 491)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 492
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__187(conn)
CALL DB.Execute(sql, 492)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 493
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__188(conn)
CALL DB.Execute(sql, 493)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 494
'...........................................................................................
sql = Aggiornamento__MEMO2__5(conn)
CALL DB.Execute(sql, 494)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__5(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 495
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__189(conn)
CALL DB.Execute(sql, 495)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 496
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__190(conn)
CALL DB.Execute(sql, 496)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 497
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__191(conn)
CALL DB.Execute(sql, 497)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(497)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 498
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__192(conn)
CALL DB.Execute(sql, 498)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(498)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 499
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__193(conn)
CALL DB.Execute(sql, 499)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(499)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 500
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__194(conn)
CALL DB.Execute(sql, 500)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__194(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 501
'...........................................................................................
sql = Aggiornamento__BOOKING3__37(conn)
CALL DB.Execute(sql, 501)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(501)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 502
'...........................................................................................
sql = Aggiornamento__BOOKING3__38(conn)
CALL DB.Execute(sql, 502)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__BOOKING3__38(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 503
'...........................................................................................
sql = Aggiornamento__REALESTATE__67(conn)
CALL DB.Execute(sql, 503)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(503)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 504
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__195(conn)
CALL DB.Execute(sql, 504)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(504)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 505
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__196(conn)
CALL DB.Execute(sql, 505)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(505)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 506
'...........................................................................................
sql = Aggiornamento__REALESTATE__68(conn)
CALL DB.Execute(sql, 506)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 507
'...........................................................................................
sql = Aggiornamento__REALESTATE__69(conn)
CALL DB.Execute(sql, 507)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 508
'...........................................................................................
sql = Aggiornamento__REALESTATE__70(conn)
CALL DB.Execute(sql, 508)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 509
'...........................................................................................
sql = Aggiornamento__REALESTATE__71(conn)
CALL DB.Execute(sql, 509)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 510
'...........................................................................................
sql = Aggiornamento__REALESTATE__72(conn)
CALL DB.Execute(sql, 510)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(510)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 511
'...........................................................................................
sql = Aggiornamento__MEMO2__6(conn)
CALL DB.Execute(sql, 511)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 512
'...........................................................................................
sql = Aggiornamento__MEMO2__7(conn)
CALL DB.Execute(sql, 512)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__7(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 513
'...........................................................................................
sql = Aggiornamento__MEMO2__8(conn)
CALL DB.Execute(sql, 513)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__8(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 514
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__197(conn)
CALL DB.Execute(sql, 514)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 515
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__198(conn)
CALL DB.Execute(sql, 515)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 516
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__199(conn)
CALL DB.Execute(sql, 516)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 517
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__200(conn)
CALL DB.Execute(sql, 517)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 518
'...........................................................................................
sql = Aggiornamento__MEMO2__9(conn)
CALL DB.Execute(sql, 518)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 519
'...........................................................................................
sql = Aggiornamento__MEMO2__10(conn)
CALL DB.Execute(sql, 519)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 520
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__201(conn)
CALL DB.Execute(sql, 520)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 521
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "ru")
CALL DB.Execute(sql, 521)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 522
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "pt")
CALL DB.Execute(sql, 522)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 523
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "cn")
CALL DB.Execute(sql, 523)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 524
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__203(conn)
CALL DB.Execute(sql, 524)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__203(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 525
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__204(conn)
CALL DB.Execute(sql, 525)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__204(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 526
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__205(conn)
CALL DB.Execute(sql, 526)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 527
'...........................................................................................
sql = Aggiornamento__BOOKING3__39(conn, "ru")
CALL DB.Execute(sql, 527)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(527)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 528
'...........................................................................................
sql = Aggiornamento__BOOKING3__39(conn, "cn")
CALL DB.Execute(sql, 528)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(528)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 529
'...........................................................................................
sql = Aggiornamento__BOOKING3__39(conn, "pt")
CALL DB.Execute(sql, 529)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(529)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 530
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__206(conn)
CALL DB.Execute(sql, 530)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(530)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 531
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__207(conn)
CALL DB.Execute(sql, 531)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(531)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 532
'...........................................................................................
sql = Aggiornamento__GUESTBOOK__2(conn)
CALL DB.Execute(sql, 532)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 533
'...........................................................................................
sql = Aggiornamento__GUESTBOOK__3(conn)
CALL DB.Execute(sql, 533)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__GUESTBOOK__3(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 534
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__208(conn)
CALL DB.Execute(sql, 534)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 535
'...........................................................................................
sql = Aggiornamento__MEMO2__11(conn)
CALL DB.Execute(sql, 535)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 536
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__209(conn)
CALL DB.Execute(sql, 536)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__209(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(536)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 537
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__210(conn)
CALL DB.Execute(sql, 537)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 538
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__211(conn)
CALL DB.Execute(sql, 538)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(538)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 539
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__212(conn)
CALL DB.Execute(sql, 539)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 540
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__213(conn)
CALL DB.Execute(sql, 540)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 541
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__214(conn)
CALL DB.ProtectedExecuteRebuild(sql, 541, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 542
'...........................................................................................
sql = Aggiornamento__BOOKING3__40(conn)
CALL DB.ProtectedExecuteRebuild(sql, 542, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 543
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__215(conn)
CALL DB.ProtectedExecuteRebuild(sql, 543, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 544
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__216(conn)
CALL DB.ProtectedExecuteRebuild(sql, 544, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 545
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__217(conn)
CALL DB.Execute(sql, 545)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__217(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 546
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__218(conn)
CALL DB.ProtectedExecuteRebuild(sql, 546, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 547
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__219(conn)
CALL DB.ProtectedExecuteRebuild(sql, 547, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 548
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__220(conn)
CALL DB.ProtectedExecuteRebuild(sql, 548, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 549
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__221(conn)
CALL DB.ProtectedExecuteRebuild(sql, 549, false, true)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 550
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__222(conn)
CALL DB.Execute(sql, 550)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 551
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__223(conn)
CALL DB.ProtectedExecuteRebuild(sql, 551, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 552
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__224(conn)
CALL DB.Execute(sql, 552)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__224(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 553
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__225(conn)
CALL DB.ProtectedExecuteRebuild(sql, 553, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 554
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__226(conn)
CALL DB.ProtectedExecuteRebuild(sql, 554, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 555
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__227(conn)
CALL DB.ProtectedExecuteRebuild(sql, 555, false, true)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 556
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__228(conn)
CALL DB.ProtectedExecuteRebuild(sql, 556, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 557
'...........................................................................................
sql = Aggiornamento__REALESTATE__73(conn)
CALL DB.Execute(sql, 557)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__73(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 558
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__229(conn)
CALL DB.ProtectedExecuteRebuild(sql, 558, false, true)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 559
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__230(conn)
CALL DB.Execute(sql, 559)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__230(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 560
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__231(conn)
CALL DB.ProtectedExecuteRebuild(sql, 560, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 561
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__232(conn)
CALL DB.ProtectedExecuteRebuild(sql, 561, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 562
'...........................................................................................
sql = Aggiornamento__MEMO2__12(conn)
CALL DB.Execute(sql, 562)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 563
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__233(conn)
CALL DB.ProtectedExecuteRebuild(sql, 563, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 564
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__234(conn)
CALL DB.ProtectedExecuteRebuild(sql, 564, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 565
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__235(conn)
CALL DB.ProtectedExecuteRebuild(sql, 565, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 566
'...........................................................................................
sql = Aggiornamento__REALESTATE__74(conn)
CALL DB.ProtectedExecuteRebuild(sql, 566, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 567
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__236(conn)
CALL DB.ProtectedExecuteRebuild(sql, 567, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 568
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__237(conn)
CALL DB.ProtectedExecuteRebuild(sql, 568, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 569
'...........................................................................................
sql = Aggiornamento__MEMO2__13(conn)
CALL DB.Execute(sql, 569)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 570
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__238(conn)
CALL DB.ProtectedExecuteRebuild(sql, 570, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 571
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__239(conn)
CALL DB.ProtectedExecuteRebuild(sql, 571, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 572
'...........................................................................................
sql = Aggiornamento__REALESTATE__75(conn)
CALL DB.Execute(sql, 572)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__75(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 573
'...........................................................................................
sql = Aggiornamento__REALESTATE__76(conn)
CALL DB.ProtectedExecuteRebuild(sql, 573, false, true)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransactionAlways()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 574
'...........................................................................................
sql = Aggiornamento__REALESTATE__77(conn)
CALL DB.ProtectedExecuteRebuild(sql, 574, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 575
'...........................................................................................
sql = Aggiornamento__REALESTATE__78(conn)
CALL DB.ProtectedExecuteRebuild(sql, 575, false, true)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 576
'...........................................................................................
sql = Aggiornamento__REALESTATE__79(conn)
CALL DB.Execute(sql, 576)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__REALESTATE__79(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 577
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__240(conn)
CALL DB.ProtectedExecuteRebuild(sql, 577, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 578
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__241(conn)
CALL DB.ProtectedExecuteRebuild(sql, 578, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 579
'...........................................................................................
sql = Aggiornamento__MEMO2__14(conn)
CALL DB.Execute(sql, 579)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 580
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__242(conn)
CALL DB.ProtectedExecuteRebuild(sql, 580, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 581
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__243(conn)
CALL DB.ProtectedExecuteRebuild(sql, 581, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 582
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__244(conn)
CALL DB.ProtectedExecuteRebuild(sql, 582, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 583
'...........................................................................................
sql = Aggiornamento__REALESTATE__81(conn)
CALL DB.ProtectedExecuteRebuild(sql, 583, false, true)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... 
%>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<!--#INCLUDE FILE="../Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../IndexContent/Tools_IndexContent.asp" -->
<% Server.ScriptTimeout=2000000 %>
<%
'...........................................................................................
set index.conn = conn
'...........................................................................................


'*******************************************************************************************
'TUTTI GLI AGGIORNAMENTI PRECEDENTI SONO ESEGUITI SU FILE SEPARATI DI SQL SERVER
'N.B.: GLI AGGIORNAMENTI DEL FRAMEWORKCORE 45 - 46 AGGIUNGONO LA TABELLA CARNET CHE GIA EISTE
'*******************************************************************************************


'*******************************************************************************************
'Esecuzione spostata in aggiornamento 32 per disallineamento versioni
'*******************************************************************************************
'AGGIORNAMENTO 27
'...........................................................................................
'modifica permessi di accesso all'area di amministrazione degli utenti 
'...........................................................................................
'sql = " UPDATE tb_siti SET sito_p1='PASS_ADMINISTRATOR', " & _
'	  " sito_p2='PASS_WEBMASTERS_ADMIN', sito_p3='PASS_USERS_ADMIN', " &_
'	  " sito_nome='next-Passport - Gestione utenti', sito_dir='../nextPassport'" &_
'	  " WHERE sito_corrente=1"
'CALL DB.Execute(sql, 27)
'*******************************************************************************************


'*******************************************************************************************
'Esecuzione spostata in aggiornamento 32 per disallineamento versioni
'*******************************************************************************************
'AGGIORNAMENTO 28
'...........................................................................................
'aggiunge campo su tabella siti che indica se l'applicazione e' un'area riservata esterna
'...........................................................................................
'sql = "ALTER TABLE tb_siti ADD sito_amministrazione bit ; " &_
'	  "UPDATE tb_siti SET sito_amministrazione=1"
'CALL DB.Execute(sql, 28)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 29
'...........................................................................................
'toglie campo sito_corrente da tabella siti del next-passport
'...........................................................................................
sql = "ALTER TABLE tb_siti DROP COLUMN sito_corrente; "
CALL DB.Execute(sql, 29)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 30
'...........................................................................................
'toglie campo sito_corrente da tabella siti del next-passport
'...........................................................................................
sql = "ALTER TABLE Stru_Ric ADD " &_
	  "		le_note_spa NTEXT, " &_
	  "		le_note_fra NTEXT, " &_
	  "		le_note_ted NTEXT "
CALL DB.Execute(sql, 30)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 31
'...........................................................................................
'aggiunge tabella per la registrazione degli accessi degli utenti
'...........................................................................................
sql = " CREATE TABLE dbo.log_admin (" &_
	  "		log_id int NOT NULL IDENTITY(1,1), " &_
	  "		log_admin_id int NULL , " &_
	  "		log_sito_id int NULL ," &_
	  "		log_data smalldatetime NULL ," &_
	  "		log_username nvarchar (50) NULL ); " &_
	  " ALTER TABLE log_admin ADD CONSTRAINT FK_log_admin__tb_admin " & _
	  " FOREIGN KEY (log_admin_id) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 31)
'*******************************************************************************************


'*******************************************************************************************
'comprende anche aggiornamenti 27 e 28 per disallineamento versioni tra versione di partenza 
'e versione di creazione degli script
'*******************************************************************************************
'AGGIORNAMENTO 32
'...........................................................................................
'aggiunge applicazione per gestione nuova area amministrativa ed aggiorna permessi
'...........................................................................................
sql = " UPDATE tb_siti SET sito_p1='PASS_ADMINISTRATOR', " & _
	  " sito_p2='PASS_WEBMASTERS_ADMIN', sito_p3='PASS_USERS_ADMIN', " &_
	  " sito_nome='next-Passport - Gestione utenti', sito_dir='../nextPassport'" &_
	  " WHERE id_sito=15; " &_
	  " ALTER TABLE tb_siti ADD sito_amministrazione bit ; " &_
	  " UPDATE tb_siti SET sito_amministrazione=1" &_
	  " SET IDENTITY_INSERT tb_siti ON " &_
	  " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (101, 'APT - Amministrazione dati dei portali', '../AptAdmin', 1, 'PORTAL_USER') " & _
	  " UPDATE rel_admin_sito SET sito_id=101 WHERE sito_id=1 " &_
	  " SET IDENTITY_INSERT tb_siti OFF "
CALL DB.Execute(sql, 32)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 33
'...........................................................................................
'toglie incremento automatico tabella siti e modifica relazioni
'...........................................................................................
sql = " ALTER TABLE tb_siti ADD id_sito_tmp integer; " &_
	  " UPDATE tb_siti SET id_sito_tmp=id_sito; " &_
	  " ALTER TABLE dbo.rel_admin_sito DROP CONSTRAINT FK_rel_admin_sito_tb_siti; " &_
	  " ALTER TABLE tb_siti DROP CONSTRAINT PK_tb_siti; " &_
	  " ALTER TABLE tb_siti DROP COLUMN id_sito; " &_
	  " ALTER TABLE tb_siti ADD id_sito INT ; " &_
	  " UPDATE tb_siti SET id_sito=id_sito_tmp; " &_
	  " ALTER TABLE tb_siti ALTER COLUMN id_sito INT NOT NULL; " &_
	  " ALTER TABLE dbo.tb_siti ADD CONSTRAINT PK_tb_siti PRIMARY KEY  NONCLUSTERED (id_sito) ; " &_
	  " ALTER TABLE tb_siti DROP COLUMN id_sito_tmp; " &_
	  " ALTER TABLE rel_admin_sito ADD CONSTRAINT FK_rel_admin_sito__tb_siti " & _
	  " FOREIGN KEY (sito_id) REFERENCES tb_siti(id_sito) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;" &_
	  " ALTER TABLE log_admin ADD CONSTRAINT FK_log_admin__tb_siti " & _
	  " FOREIGN KEY (log_sito_id) REFERENCES tb_siti(id_sito) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 33)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 34
'...........................................................................................
'aggiorna relazioni tra tabella tb_siti e rel_admin_suto
'...........................................................................................
sql = " ALTER TABLE rel_admin_sito DROP CONSTRAINT FK_rel_admin_sito_tb_admin; " &_
	  " ALTER TABLE rel_admin_sito ADD CONSTRAINT FK_rel_admin_sito__tb_admin " & _
	  " FOREIGN KEY (admin_id) REFERENCES tb_admin(id_admin) " &_
	  " ON UPDATE CASCADE ON DELETE CASCADE;"
CALL DB.Execute(sql, 34)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 35
'...........................................................................................
'cambia indici per applicazione next-passport e trasporta permessi nella nuova applicazione
'...........................................................................................
sql = " UPDATE tb_siti SET sito_p1='PASS_ADMINISTRATOR', " & _
	  " sito_p2='PASS_WEBMASTERS_ADMIN', sito_p3='PASS_USERS_ADMIN', " &_
	  " sito_nome='next-Passport - Gestione utenti', sito_dir='../nextPassport'" &_
	  " WHERE id_sito=1; " &_
	  " UPDATE rel_admin_sito SET sito_id=1 WHERE sito_id=15; " &_
	  " DELETE tb_siti WHERE id_sito=15;"
CALL DB.Execute(sql, 35)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 36
'...........................................................................................
'aggiunge tabelle per la gestione degli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE dbo.tb_Utenti (" & _
	  " 	ut_ID INT IDENTITY(1,1), " & _
	  " 	ut_NextCom_ID INT NOT NULL, " & _
	  " 	ut_login nvarchar(50) NULL, " & _
	  " 	ut_password nvarchar(50) NULL, " & _
	  " 	ut_Abilitato bit, " & _
	  " 	ut_ScadenzaAccesso SMALLDATETIME NULL ) ; " &_
	  " ALTER TABLE dbo.tb_Utenti ADD CONSTRAINT PK_tb_Utenti PRIMARY KEY  NONCLUSTERED (ut_ID) "
CALL DB.Execute(sql, 36)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 37
'...........................................................................................
'aggiunge tabelle per la gestione dei permessi degli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE dbo.rel_utenti_sito (" & _
	  "		rel_id INT IDENTITY(1,1), " & _
	  "		rel_ut_id INT NOT NULL, " & _
	  "		rel_sito_id INT NOT NULL, " & _
	  "		rel_permesso INT NOT NULL ); " &_
	  " ALTER TABLE dbo.rel_utenti_sito ADD CONSTRAINT PK_rel_utenti_sito PRIMARY KEY  NONCLUSTERED (rel_ID); " &_
	  " ALTER TABLE rel_utenti_sito ADD CONSTRAINT FK_rel_utenti_sito__tb_utenti " &_
   	  " 	FOREIGN KEY (rel_ut_id) REFERENCES tb_Utenti (Ut_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_	
	  " ALTER TABLE rel_utenti_sito ADD CONSTRAINT FK_rel_utenti_sito__tb_siti " &_
   	  " 	FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 37)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 38
'...........................................................................................
'aggiunge campo su tabella siti che indica se l'applicazione e' un'area riservata esterna
'...........................................................................................
sql = "ALTER TABLE tb_siti ADD sito_rubrica_area_riservata INT NULL"
CALL DB.Execute(sql, 38)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 39
'...........................................................................................
'cancella tabelle non piu' usate nel database: cancella tabelle per mappe berenice
'...........................................................................................
sql = " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('LayersMappe') AND sysstat & 0xf = 3) " &_
	  " DROP TABLE LayersMappe; " &_
	  " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('LayerZonaDettagli') AND sysstat & 0xf = 3) " &_
	  " DROP TABLE LayerZonaDettagli "
CALL DB.Execute(sql, 39)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 40
'...........................................................................................
'cancella tabelle non piu' usate nel database: cancella tabella daVedere per itinerari
'...........................................................................................
sql = " ALTER TABLE dbo.daVedere DROP CONSTRAINT PK__daVedere__137DBFF6; " &_
	  " ALTER TABLE dbo.daVedere DROP CONSTRAINT FK__daVedere__idLuog__681373AD; " &_
	  " ALTER TABLE dbo.daVedere DROP CONSTRAINT FK_daVedere_Tappe; " &_
	  " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.delete_ALL_Itinerari') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.delete_ALL_Itinerari; " &_
	  " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.delete_ALL_luoghi') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.delete_ALL_luoghi; " &_
	  " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.delete_itinerario') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.delete_itinerario; " &_
	  " ALTER PROCEDURE dbo.Delete_Luoghi ( @ID int) AS " &_
	  " DELETE FROM Imag_lu WHERE id_Luogo = @ID " &_
	  " DELETE FROM doveAccade WHERE id_luogo = @ID " &_
	  " DELETE FROM Luoghi WHERE ID = @ID ;  " &_
	  " ALTER  PROCEDURE dbo.Delete_TipoLuoghi ( @IDL int) AS " &_
	  " DELETE FROM Imag_lu WHERE id_Luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " &_
	  " DELETE FROM doveAccade WHERE id_luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " &_
	  " DELETE FROM Luoghi WHERE ID IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " &_
	  " DELETE FROM TipoLuoghi WHERE IDL = @IDL; " &_
	  " DROP TABLE daVedere "
CALL DB.Execute(sql, 40)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 41
'...........................................................................................
'cancella tabelle non piu' usate nel database: cancella tabella tappe per itinerari
'...........................................................................................
sql = " ALTER TABLE dbo.Tappe DROP CONSTRAINT FK_Tappe_Itinerari; " &_
	  " ALTER TABLE dbo.Tappe DROP CONSTRAINT PK_Tappe; " &_
	  " DROP TABLE tappe"
CALL DB.Execute(sql, 41)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 42
'...........................................................................................
'cancella tabelle non piu' usate nel database: cancella tabella itinerari
'...........................................................................................
sql = " ALTER TABLE dbo.Itinerari DROP CONSTRAINT FK__Itinerari__zona__236943A5; " &_
	  " ALTER TABLE dbo.Itinerari DROP CONSTRAINT PK__Itinerari__7310F064; " &_
	  " ALTER PROCEDURE dbo.Delete_SubZona (@SubZona_ID int ) AS " &_
	  " UPDATE Not_Util SET subZona = '' WHERE subZona = @SubZona_ID " &_
	  " UPDATE Stru_Ric SET rif_subzona='' WHERE rif_subZona = @SubZona_ID " &_
	  " UPDATE Luoghi SET SubZona='' WHERE SubZona = @SubZona_ID " &_
	  " DELETE FROM SubZone WHERE id_subzona = @SubZona_ID ; " &_
	  " DROP TABLE dbo.itinerari"
CALL DB.Execute(sql, 42)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 43
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella relCliCat
'...........................................................................................
sql = " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.Delete_Categoria') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.Delete_Categoria; " &_
	  " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.Delete_Contatto') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.Delete_Contatto; " &_
	  " ALTER TABLE dbo.relCliCat DROP CONSTRAINT PK__relCliCat__7AB2122C; " &_
	  " ALTER TABLE dbo.relCliCat DROP CONSTRAINT FK__relCliCat__IDCli__489AC854; " &_
	  " ALTER TABLE dbo.relCliCat DROP CONSTRAINT FK__relCliCat__IDCat__32F66B4F; " &_
	  " DROP TABLE relCliCat "
CALL DB.Execute(sql, 43)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 44
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_cli_logs
'...........................................................................................
sql = " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.Delete_Email') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.Delete_Email; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT FK__tb_cli_lo__tb_ed__42E1EEFE; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT FK__tb_cli_lo__tb_em__3D73F9C2; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT PK__tb_cli_logs__7F76C749; " &_
	  " DROP TABLE tb_cli_logs "
CALL DB.Execute(sql, 44)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 44
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_cli_logs
'...........................................................................................
sql = " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.Delete_Email') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.Delete_Email; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT FK__tb_cli_lo__tb_ed__42E1EEFE; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT FK__tb_cli_lo__tb_em__3D73F9C2; " &_
	  " ALTER TABLE dbo.tb_cli_logs DROP CONSTRAINT PK__tb_cli_logs__7F76C749; " &_
	  " DROP TABLE tb_cli_logs "
CALL DB.Execute(sql, 44)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 45
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tbCatClienti
'...........................................................................................
sql = " ALTER TABLE dbo.tbCatClienti DROP CONSTRAINT PK__tbCatClienti__62DA889B; " &_
	  " DROP TABLE tbCatClienti " 
CALL DB.Execute(sql, 45)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 46
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_clienti
'...........................................................................................
sql = " ALTER TABLE dbo.tb_clienti DROP CONSTRAINT PK__tb_clienti__40058253; " &_
	  " DROP TABLE tb_clienti " 
CALL DB.Execute(sql, 46)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 47
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_emails
'...........................................................................................
sql = " ALTER TABLE dbo.tb_emails DROP CONSTRAINT PK__tb_emails__65B6F546; " &_
	  " DROP TABLE tb_emails " 
CALL DB.Execute(sql, 47)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 48
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_clienti
'...........................................................................................
sql = " ALTER TABLE dbo.tb_links DROP CONSTRAINT PK_tb_links; " &_
	  " DROP TABLE tb_links " 
CALL DB.Execute(sql, 48)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 49
'...........................................................................................
'cancella stored procedure non piu' usate
'...........................................................................................
sql = " IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('dbo.Delete_Sito') AND sysstat & 0xf = 4) " &_
	  " DROP PROCEDURE dbo.Delete_Sito" 
CALL DB.Execute(sql, 49)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 50
'...........................................................................................
'cancella tabelle non piu' usate nel database: tabella tb_Amministratori
'...........................................................................................
sql = " ALTER TABLE dbo.tb_amministratori DROP CONSTRAINT PK_tb_amministratori; " &_
	  " DROP TABLE tb_amministratori " 
CALL DB.Execute(sql, 50)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 51
'...........................................................................................
'aggiunge tabelle per indirizzario
'...........................................................................................
sql = " CREATE TABLE dbo.log_cnt_email (" &_
	  " 	log_id INT IDENTITY(1,1) NOT NULL ," &_
	  " 	log_cnt_id int NULL ," &_
	  " 	log_email nvarchar(50) NULL ," &_
	  " 	log_email_id int NULL ); " &_
	  " ALTER TABLE dbo.log_cnt_email ADD CONSTRAINT PK_log_cnt_email PRIMARY KEY NONCLUSTERED (log_ID); " &_
	  " CREATE TABLE dbo.rel_dip_email (" &_
	  " 	rel_id INT IDENTITY(1,1) NOT NULL ," &_
	  " 	rel_emailSender nvarchar(250) NOT NULL ," &_
	  "		rel_emailSenderID INT NULL ," &_
	  " 	rel_emailID INT NULL ," &_
	  "		rel_dipID INT NULL , " &_
	  " 	rel_Read BIT NOT NULL , " &_
	  " 	rel_Reply BIT NOT NULL ); " &_
	  " ALTER TABLE dbo.rel_dip_email ADD CONSTRAINT PK_rel_dip_email PRIMARY KEY NONCLUSTERED (rel_id); " &_
	  " CREATE TABLE dbo.rel_rub_ind (" &_
	  "		id_rub_ind INT IDENTITY(1,1) NOT NULL , " &_
	  "		id_indirizzo INT NULL , " &_
	  "		id_rubrica INT NULL ); " &_
	  " ALTER TABLE dbo.rel_rub_ind ADD CONSTRAINT PK_rel_rub_ind PRIMARY KEY NONCLUSTERED (id_rub_ind); "  &_ 
	  " CREATE TABLE dbo.tb_Indirizzario (" &_
	  "		IDElencoIndirizzi INT IDENTITY(1,1) NOT NULL ," &_
	  "		NomeElencoIndirizzi nvarchar(100) NULL ," &_
	  "		SecondoNomeElencoIndirizzi nvarchar(30) NULL ," &_
	  "		CognomeElencoIndirizzi nvarchar(100) NULL ," &_
	  "		TitoloElencoIndirizzi nvarchar(50) NULL ," &_
	  "		NomeOrganizzazioneElencoIndirizzi nvarchar(255) NULL ," &_
	  "		QualificaElencoIndirizzi nvarchar(250) NULL ," &_
	  "		IndirizzoElencoIndirizzi nvarchar(255) NULL ," &_
	  "		CittaElencoIndirizzi nvarchar(50) NULL ," &_
	  "		StatoProvElencoIndirizzi nvarchar(50) NULL ," &_
	  "		ZonaElencoIndirizzi nvarchar(50) NULL ," &_
	  "		CAPElencoIndirizzi nvarchar(20) NULL ," &_
	  "		CountryElencoIndirizzi nvarchar(50) NULL ," &_
	  "		DTNASCElencoIndirizzi SMALLDATETIME NULL ," &_
	  "		NoteElencoIndirizzi ntext NULL ," &_
	  "		isSocieta bit NOT NULL ," &_
	  "		ModoRegistra nvarchar(255) NOT NULL ," &_
	  "		DataIscrizione smalldatetime NULL ," &_
	  "		LockedByApplication int NULL ," &_
	  "		ApplicationsLocker nvarchar(50) NULL ); " &_
	  "	ALTER TABLE dbo.tb_Indirizzario ADD CONSTRAINT PK_tb_Indirizzario PRIMARY KEY NONCLUSTERED (IDElencoIndirizzi); " &_
	  "	CREATE TABLE dbo.tb_ValoriNumeri ( " &_
	  "		 id_ValoreNumero INT IDENTITY(1,1) NOT NULL , " &_
	  "		 id_Indirizzario int NULL , " &_
	  "		 id_TipoNumero int NULL , " &_
	  "		 ValoreNumero nvarchar (50) NULL , " &_
	  "		 email_default bit NOT NULL ); " &_
	  "	ALTER TABLE dbo.tb_ValoriNumeri ADD CONSTRAINT PK_tb_ValoriNumeri PRIMARY KEY NONCLUSTERED (id_ValoreNumero); " &_
	  " CREATE TABLE dbo.tb_email ( " &_
	  "		 email_id INT IDENTITY(1,1) NOT NULL , " &_
	  "		 email_text ntext NULL , " &_
	  "		 email_object nvarchar(200) NULL , " &_
	  "		 email_data SMALLDATETIME NULL , " &_
	  "		 email_dipgenera int NULL , " &_
	  "		 email_docs ntext NULL , " &_
	  "		 email_page_ID int NULL , " &_
	  "		 email_page_owned bit NOT NULL , " &_
	  "		 email_in bit NOT NULL , " &_
	  "		 email_MessageID nvarchar (100) NULL , " &_
	  "		 email_UIDL int NOT NULL , " &_
	  "		 email_Account int NOT NULL , " &_
	  "		 email_To nvarchar (250) NULL , " &_
	  "		 email_CC nvarchar (250) NULL);  " &_
	  "	ALTER TABLE dbo.tb_email ADD CONSTRAINT PK_tb_email PRIMARY KEY NONCLUSTERED (email_id); " &_
	  " CREATE TABLE dbo.tb_emailConfig ( " &_
	  "		 config_id int IDENTITY(1, 1) NOT NULL , " &_
	  "		 config_host nvarchar(250) NOT NULL , " &_
	  "		 config_port int NULL , " &_
	  "		 config_user nvarchar(50) NOT NULL , " &_
	  "		 config_pass nvarchar(50) NOT NULL , " &_
	  "		 config_protocol nvarchar(5) NOT NULL , " &_
	  "		 config_email nvarchar(250) NOT NULL , " &_
	  "		 config_deleteMessage bit NOT NULL , " &_
	  "		 config_delayDelMessage int NULL , " &_
	  "		 config_id_empl int NOT NULL ); " &_
	  "	ALTER TABLE dbo.tb_emailConfig ADD CONSTRAINT PK_tb_emailConfig PRIMARY KEY NONCLUSTERED (config_id); " &_
	  " CREATE TABLE dbo.tb_gruppi (" &_
	  "		 id_Gruppo INT IDENTITY(1,1) NOT NULL , " &_
	  "		 nome_Gruppo nvarchar(50) NULL ); " &_
	  "	ALTER TABLE dbo.tb_gruppi ADD CONSTRAINT PK_tb_gruppi PRIMARY KEY NONCLUSTERED (id_Gruppo); " &_
	  " CREATE TABLE dbo.tb_rel_dipgruppi ( " &_
	  "		 id_rel_dipgruppi int IDENTITY(1,1) NOT NULL , " &_
	  "		 id_impiegato int NULL , " &_
	  "		 id_gruppo int NULL ) ; " &_
	  "	ALTER TABLE dbo.tb_rel_dipgruppi ADD CONSTRAINT PK_tb_rel_dipgruppi PRIMARY KEY NONCLUSTERED (id_rel_dipgruppi); " &_
	  " CREATE TABLE dbo.tb_rel_gruppirubriche ( " &_
	  "		 id_rel_grupprub INT IDENTITY(1,1) NOT NULL , " &_
	  "		 id_dellaRubrica int NULL , " &_
	  "		 id_Gruppo_assegnato int NULL ); " &_
	  "	ALTER TABLE dbo.tb_rel_gruppirubriche ADD CONSTRAINT PK_tb_rel_gruppirubriche PRIMARY KEY NONCLUSTERED (id_rel_grupprub); " &_
	  " CREATE TABLE dbo.tb_rubriche ( " &_
	  "		 id_Rubrica INT IDENTITY NOT NULL ," &_
	  "		 nome_Rubrica nvarchar(250) NULL ," &_
	  "		 note_Rubrica ntext NULL ," &_
	  "		 locked_rubrica bit NOT NULL ," &_
	  "		 rubrica_esterna bit NOT NULL ); " &_
	  "	ALTER TABLE dbo.tb_rubriche ADD CONSTRAINT PK_tb_rubriche PRIMARY KEY NONCLUSTERED (id_Rubrica); " &_
	  " CREATE TABLE dbo.tb_tipNumeri ( " &_
	  "		 id_tipoNumero INT NOT NULL ," &_
	  "		 nome_tipoNumero nvarchar(250) NULL ," &_
	  "		 tipoNumero nvarchar(1) NULL ); " &_
	  "	ALTER TABLE dbo.tb_tipNumeri ADD CONSTRAINT PK_tb_tipNumeri PRIMARY KEY NONCLUSTERED (id_tipoNumero); " &_
	  " INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (1, 'Telefono'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (2, 'Telefono Ufficio'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (3, 'Telefono Cellulare'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (4, 'Telefono Casa'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (5, 'Numero Fax'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (6, 'Indirizzo Email'); " &_
	  "	INSERT INTO tb_tipNumeri( id_tipoNumero, nome_tipoNumero) VALUES (7, 'Indirizzo Web') "
CALL DB.Execute(sql, 51)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 52
'...........................................................................................
'aggiunge relazioni indirizzario
'...........................................................................................
sql = " ALTER TABLE dbo.tb_ValoriNumeri ADD CONSTRAINT FK_tb_ValoriNumeri_tb_tipNumeri " &_
	  " 	FOREIGN KEY (id_TipoNumero) REFERENCES tb_tipNumeri (id_tipoNumero) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.tb_ValoriNumeri ADD CONSTRAINT FK_tb_ValoriNumeri_tb_Indirizzario " &_
	  "		FOREIGN KEY (id_Indirizzario) REFERENCES tb_Indirizzario (IDElencoIndirizzi) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.rel_rub_ind ADD CONSTRAINT FK_rel_rub_ind_tb_Indirizzario " &_
	  " 	FOREIGN KEY (id_indirizzo) REFERENCES tb_Indirizzario (IDElencoIndirizzi) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.rel_rub_ind ADD CONSTRAINT FK_rel_rub_ind_tb_rubriche " &_
	  "		FOREIGN KEY (id_rubrica) REFERENCES tb_rubriche (id_Rubrica) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.tb_rel_gruppirubriche ADD CONSTRAINT FK_tb_rel_gruppirubriche_tb_rubriche  " &_
	  "		FOREIGN KEY (id_dellaRubrica) REFERENCES tb_rubriche (id_Rubrica) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  "	ALTER TABLE dbo.tb_rel_gruppirubriche ADD CONSTRAINT FK_tb_rel_gruppirubriche_tb_gruppi " &_
	  "		FOREIGN KEY (id_Gruppo_assegnato) REFERENCES tb_gruppi (id_Gruppo) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  "	ALTER TABLE dbo.tb_rel_dipgruppi ADD CONSTRAINT FK_tb_rel_dipgruppi_tb_gruppi " &_
	  "		FOREIGN KEY (id_gruppo) REFERENCES tb_gruppi (id_Gruppo) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  "		ALTER TABLE dbo.tb_rel_dipgruppi ADD CONSTRAINT FK_tb_rel_dipgruppi_tb_admin " &_
	  "		FOREIGN KEY (id_impiegato) REFERENCES tb_admin (id_admin) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.tb_emailConfig ADD CONSTRAINT FK_tb_emailConfig_tb_admin " &_ 
	  " 	FOREIGN KEY (config_id_empl) REFERENCES tb_admin (id_admin) " &_ 
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.log_cnt_email ADD CONSTRAINT FK_log_cnt_email_tb_Indirizzario " &_
	  "		FOREIGN KEY (log_cnt_id) REFERENCES tb_Indirizzario (IDElencoIndirizzi) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.log_cnt_email ADD CONSTRAINT FK_log_cnt_email_tb_email " &_
	  "		FOREIGN KEY (log_email_id) REFERENCES tb_email (email_id) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.tb_Utenti ADD CONSTRAINT FK_tb_Utenti_tb_Indirizzario " &_
	  " 	FOREIGN KEY (ut_NextCom_ID) REFERENCES tb_Indirizzario (IDElencoIndirizzi) " &_
	  "		ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.rel_dip_email ADD CONSTRAINT FK_rel_dip_email_tb_email " &_
	  " 	FOREIGN KEY (rel_emailID) REFERENCES tb_email (email_id) " &_
	  " 	ON DELETE CASCADE  ON UPDATE CASCADE; " &_
	  " ALTER TABLE dbo.rel_dip_email ADD CONSTRAINT FK_rel_dip_email_tb_admin " &_
	  " 	FOREIGN KEY (rel_dipID) REFERENCES tb_admin (id_admin) " &_
	  " 	ON DELETE CASCADE  ON UPDATE CASCADE;  " &_
	  " ALTER TABLE dbo.tb_siti WITH NOCHECK ADD CONSTRAINT FK_tb_siti_tb_rubriche " &_
	  "		FOREIGN KEY (sito_rubrica_area_riservata) REFERENCES tb_rubriche (id_Rubrica)" &_
	  "		NOT FOR REPLICATION; " &_
	  "	ALTER TABLE dbo.tb_siti NOCHECK CONSTRAINT FK_tb_siti_tb_rubriche " 
CALL DB.Execute(sql, 52)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 53
'...........................................................................................
'modifica permessi di accesso all'area di amministrazione degli utenti 
'...........................................................................................
sql = "UPDATE tb_siti SET sito_p1='PASS_ADMIN', sito_p2='PASS_AMMINISTRATORI', sito_p3='PASS_UTENTI' " &_
	  " WHERE id_sito=1"
CALL DB.Execute(sql, 53)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 54
'...........................................................................................
'modifica permessi di accesso all'area di amministrazione degli utenti 
'...........................................................................................
sql = " ALTER TABLE tb_admin ALTER COLUMN admin_email nvarchar(100) NULL; " &_
	  " UPDATE tb_admin SET admin_email=RTRIM(admin_email) "
CALL DB.Execute(sql, 54)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 55
'...........................................................................................
'aggiunge tabella per log degli accessi per gli utenti dell'area riservata
'...........................................................................................
sql = " CREATE TABLE dbo.log_utenti (" & _
	  "		log_id INT IDENTITY(1,1) NOT NULL, " & _
	  "		log_ut_id INT NOT NULL, " &_
	  "		log_sito_id INT NOT NULL, " &_
	  "		log_data SMALLDATETIME NULL, " &_
	  "		log_username nvarchar(50) NULL ); " &_
	  "	ALTER TABLE dbo.log_utenti ADD CONSTRAINT PK_log_utenti PRIMARY KEY NONCLUSTERED (log_id); " &_
	  " ALTER TABLE log_utenti ADD CONSTRAINT FK_log_utenti__tb_utenti " &_
   	  " 	FOREIGN KEY (log_ut_id) REFERENCES tb_Utenti (Ut_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " &_
	  " ALTER TABLE log_utenti ADD CONSTRAINT FK_log_utenti__tb_siti " &_
   	  " 	FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " 
CALL DB.Execute(sql, 55)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 56
'...........................................................................................
'inserisce next-Com nelle applicazioni
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2) " & _
	  " VALUES (3, 'next-Com - Gestione comunicazioni', '../nextCom', 1, 'COM_ADMIN', 'COM_USER') "
CALL DB.Execute(sql, 56)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 57
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento next-web
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (102, 'APT - Archivio immagini', '../Foto/admin', 1, 'FOTO_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=102 WHERE sito_id=2; " &_
	  " UPDATE tb_siti SET sito_nome = 'next-Web - Gestione grafica e contenuti', sito_dir='../nextWeb', " &_
	  "	sito_p1='WEB_ADMIN', sito_p2='WEB_POWER_USER', sito_p3='WEB_USER' WHERE id_sito=2; " &_
	  " UPDATE rel_admin_sito SET sito_id=2 WHERE sito_id=5; " &_
	  " DELETE FROM tb_siti WHERE id_sito=5 "
CALL DB.Execute(sql, 57)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 57
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento next-web
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (102, 'APT-Foto - Archivio immagini', '../AptFoto', 1, 'FOTO_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=102 WHERE sito_id=2; " &_
	  " UPDATE tb_siti SET sito_nome = 'next-Web - Gestione grafica e contenuti', sito_dir='../nextWeb', " &_
	  "	sito_p1='WEB_ADMIN', sito_p2='WEB_POWER_USER', sito_p3='WEB_USER' WHERE id_sito=2; " &_
	  " UPDATE rel_admin_sito SET sito_id=2 WHERE sito_id=5; " &_
	  " DELETE FROM tb_siti WHERE id_sito=5 "
CALL DB.Execute(sql, 57)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 58
'...........................................................................................
'aggiunge tabella log dei download delle circolari
'...........................................................................................
sql = "	ALTER TABLE dbo.tb_circolari ADD CONSTRAINT PK_tb_circolari PRIMARY KEY NONCLUSTERED (CI_id); " &_
	  " CREATE TABLE dbo.log_circolari (" & _
	  "		log_id INT IDENTITY(1,1) NOT NULL, " & _
	  "		log_ut_id INTEGER NOT NULL, " &_
	  "		log_dip_id INTEGER NOT NULL, " &_
	  "		log_ci_id INTEGER NOT NULL, " &_
	  "		log_data DATETIME NULL ); " &_
	  "	ALTER TABLE dbo.log_circolari ADD CONSTRAINT PK_log_circolari PRIMARY KEY NONCLUSTERED (log_id); " &_
	  " ALTER TABLE log_circolari ADD CONSTRAINT FK_log_circolari__tb_Circolari " &_
   	  " 	FOREIGN KEY (log_ci_id) REFERENCES tb_circolari (CI_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " 
CALL DB.Execute(sql, 58)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 59
'...........................................................................................
'aggiunge tabella log dei download delle circolari
'...........................................................................................
sql = "	ALTER TABLE dbo.tb_controlloQ ADD CONSTRAINT PK_tb_controlloQ PRIMARY KEY NONCLUSTERED (CI_id); " &_
	  " CREATE TABLE dbo.log_controlloQ (" & _
	  "		log_id INT IDENTITY(1,1) NOT NULL, " & _
	  "		log_ut_id INTEGER NOT NULL, " &_
	  "		log_dip_id INTEGER NOT NULL, " &_
	  "		log_ci_id INTEGER NOT NULL, " &_
	  "		log_data DATETIME NULL ); " &_
	  "	ALTER TABLE dbo.log_controlloQ ADD CONSTRAINT PK_log_controlloQ PRIMARY KEY NONCLUSTERED (log_id); " &_
	  " ALTER TABLE log_controlloQ ADD CONSTRAINT FK_log_controlloQ__tb_controlloQ " &_
   	  " 	FOREIGN KEY (log_Ci_id) REFERENCES tb_controlloQ (CI_ID) " &_
	  " 	ON UPDATE CASCADE ON DELETE CASCADE; " 
CALL DB.Execute(sql, 59)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 60
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento next-Memo
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2) " & _
	  " VALUES (8, 'APT - Circolari interne', '../AptCircolari', 1, 'MEMO_ADMIN', 'MEMO_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=8 WHERE sito_id=18; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 18; "
CALL DB.Execute(sql, 60)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 61
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento Apt Controllo qualita'
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3) " & _
	  " VALUES (103, 'APT - Controllo qualita''', '../AptControlloQ', 1, 'CQ_ADMIN', 'CQ_POWEW_USER', 'CQ_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=103 WHERE sito_id=27 AND rel_as_permesso=1; " &_
	  " UPDATE rel_admin_sito SET sito_id=103, rel_as_permesso=2 WHERE sito_id=27 AND rel_as_permesso=3; " &_
	  " UPDATE rel_admin_sito SET sito_id=103, rel_as_permesso=3 WHERE sito_id=27 AND rel_as_permesso=2; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 27; "
CALL DB.Execute(sql, 61)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 62
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento Apt Controllo qualita'
'...........................................................................................
sql = " ALTER TABLE tb_Circolari ADD CI_protetto bit NULL; " &_
	  " UPDATE tb_circolari SET CI_Protetto=0; " &_
	  " ALTER TABLE tb_Circolari ALTER COLUMN CI_Protetto bit NOT NULL; " &_
	  " ALTER TABLE tb_ControlloQ ADD CI_protetto bit NULL; " &_
	  " UPDATE tb_ControlloQ SET CI_Protetto=CI_SuperUser; " &_
	  " ALTER TABLE tb_ControlloQ ALTER COLUMN CI_Protetto bit NOT NULL; " &_
	  " ALTER TABLE tb_ControlloQ DROP COLUMN CI_SuperUser " 
CALL DB.Execute(sql, 62)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 63
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi dell'applicazione circolari interne
'...........................................................................................
sql = " UPDATE tb_siti set sito_p3='MEMO_USER', sito_p2='MEMO_POWER_USER' WHERE id_sito=8; " &_
	  " UPDATE rel_admin_sito SET rel_as_permesso=3 WHERE rel_as_permesso=2 AND sito_id=8 "
CALL DB.Execute(sql, 63)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 64
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento Apt bussola
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (104, 'APT - Estrazione dati Bussola', '../AptBussola', 1, 'BUSSOLA_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=104 WHERE sito_id=24 AND rel_as_permesso=1; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 24; "
CALL DB.Execute(sql, 64)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 65
'...........................................................................................
'esegue modifiche alle applicazioni ed ai permessi per inserimento Apt bussola
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (105, 'APT - LEO, la rivista di Venezia', '../Leo/Admin', 1, 'LEO_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=105 WHERE sito_id=4 ; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 4; "
CALL DB.Execute(sql, 65)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 66
'...........................................................................................
'elimina applicazione mappe
'...........................................................................................
sql = " DELETE FROM tb_siti WHERE id_sito = 26; "
CALL DB.Execute(sql, 66)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 67
'...........................................................................................
'modifica permessi ed inserisce applicazione prenotazione eventi
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3) " & _
	  " VALUES (106, 'APT - Gestione prenotazioni eventi', '../prenotazioni', 1, 'Employee', 'Association', 'Administrator'); " &_
	  " UPDATE rel_admin_sito SET sito_id=106 WHERE sito_id=14; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 14; "
CALL DB.Execute(sql, 67)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 68
'...........................................................................................
'modifica permessi ed inserisce applicazione prenotazione VeniceCard
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2) " & _
	  " VALUES (107, 'VeniceCard - Gestione prenotazioni', '../Venicecard/admin', 1, 'VENICECARD_ADMIN', 'VENICECARD_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=107 WHERE sito_id=20; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 20; "
CALL DB.Execute(sql, 68)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 69
'...........................................................................................
'modifica permessi ed inserisce applicazione prenotazione VeneziaSi
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " & _
	  " VALUES (108, 'VeneziaSi - Gestione prenotazioni', '../PrenotazioniVeneziaSi', 1, 'PRENO_USER'); " &_
	  " UPDATE rel_admin_sito SET sito_id=108 WHERE sito_id=25; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 25; "
CALL DB.Execute(sql, 69)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 70
'...........................................................................................
'modificaproprieta' tabela prenotazioni VeneziaSi
'...........................................................................................
sql = " sp_changeobjectowner 'goveniceAdm.tb_prenotazioniVeneziaSi' , 'dbo'"
CALL DB.Execute(sql, 70)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 71
'...........................................................................................
'modifica permessi ed inserisce applicazione Gestione Villa Widmann
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_p4) " & _
	  " VALUES (109, 'APT - Gestione Villa Widmann', '../widmann', 1, 'Administrator', 'Manager', 'Employee', 'Customer' ); " &_
	  " UPDATE rel_admin_sito SET sito_id=109 WHERE sito_id=19; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 19; "
CALL DB.Execute(sql, 71)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 72
'...........................................................................................
'modifica permessi ed inserisce applicazione prenotazioni turive
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2) " & _
	  " VALUES (110, 'TU.RI.VE. - Prenotazioni alberghiere', '../prenotazioni_TURIVE', 1, 'TURIVE_ADMIN', 'TURIVE_AGENZIA'); " &_
	  " UPDATE rel_admin_sito SET sito_id=110 WHERE sito_id=22; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 22; "
CALL DB.Execute(sql, 72)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 73
'...........................................................................................
'modifica permessi ed inserisce applicazione prenotazioni unindustria
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3) " & _
	  " VALUES (111, 'UNINDUSTRIA - Prenotazioni alberghiere', '../prenotazioni_UNI', 1, 'UNI_ADMIN', 'UNI_AGENZIA', 'UNI_ALBERGO'); " &_
	  " UPDATE rel_admin_sito SET sito_id=111 WHERE sito_id=23; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 23; "
CALL DB.Execute(sql, 73)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 74
'...........................................................................................
'modifica permessi ed inserisce applicazione magazzino centrale
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_p4, sito_p5) " & _
	  " VALUES (112, 'APT - Magazzino centrale', '../Mag_centrale', 1, 'IDUtente', 'Carico', 'Magazzino', 'Ordini', 'Admin'); " &_
	  " UPDATE rel_admin_sito SET sito_id=112 WHERE sito_id=6; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 6; "
CALL DB.Execute(sql, 74)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 75
'...........................................................................................
'modifica permessi ed inserisce applicazione magazzino omaggistica
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_p4, sito_p5) " & _
	  " VALUES (113, 'APT - Magazzino omaggistica', '../Mag_Omaggistica', 1, 'IDUtente', 'Carico', 'Magazzino', 'Ordini', 'Admin'); " &_
	  " UPDATE rel_admin_sito SET sito_id=113 WHERE sito_id=21; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 21; "
CALL DB.Execute(sql, 75)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 76
'...........................................................................................
'modifica permessi ed inserisce applicazione magazzino spedizioni
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_p4, sito_p5) " & _
	  " VALUES (114, 'APT - Magazzino spedizioni', '../Mag_spedizioni', 1, 'IDUtente', 'Carico', 'Magazzino', 'Ordini', 'Admin'); " &_
	  " UPDATE rel_admin_sito SET sito_id=114 WHERE sito_id=16; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 16; "
CALL DB.Execute(sql, 76)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 77
'...........................................................................................
'modifica permessi ed inserisce applicazione presenze e riunisce applicazione del report
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_p4) " & _
	  " VALUES (115, 'APT - Presenze', '../Presenze', 1, 'Administrator', 'Admin_service', 'Employee', 'Worker'); " &_
	  " UPDATE rel_admin_sito SET sito_id=115 WHERE sito_id=11; " &_
	  " UPDATE rel_admin_sito SET sito_id=115, rel_as_permesso=4 WHERE sito_id=13; " &_
	  " DELETE FROM tb_siti WHERE id_sito = 11 OR id_sito=13; "
CALL DB.Execute(sql, 77)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 78
'...........................................................................................
'aggiunge tabella per applicazione presenze per tracciare quando viene eseguita 
'la procedura di cambio dell'anno
'...........................................................................................
sql = " CREATE TABLE dbo.tb_CambioAnno_Presenze (data_esecuzione SMALLDATETIME NULL ); " &_
	  " INSERT INTO tb_CambioAnno_Presenze(data_esecuzione) VALUES (CONVERT(DATETIME, '2004 - 01 - 01 00:00:00', 102));"
CALL DB.Execute(sql, 78)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 79
'...........................................................................................
'cancella procedura GET_ALL_DIPENDENTI perche' non piu' usata
'...........................................................................................
sql = " DROP PROCEDURE GET_ALL_DIPENDENTI"
CALL DB.Execute(sql, 79)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 80
'...........................................................................................
'aggiunge campi in lingua multipla e per gestione bussola alla tabella eventi 
'...........................................................................................
sql = " ALTER TABLE Eventi ADD " &_
	  " 	Orario_eng ntext NULL, " &_
	  " 	Orario_fra ntext NULL, " &_
	  " 	Orario_ted ntext NULL, " &_
	  " 	Orario_spa ntext NULL, " &_
	  "		Descr_luogo_eng ntext NULL, " &_
	  " 	Descr_luogo_fra ntext NULL, " &_
	  " 	Descr_luogo_ted ntext NULL, " &_
	  " 	Descr_luogo_spa ntext NULL, " &_
	  " 	Descr_Bussola_Ita ntext NULL, " &_
	  " 	Descr_Bussola_eng ntext NULL, " &_
	  " 	Descr_Bussola_fra ntext NULL, " &_
	  " 	Descr_Bussola_ted ntext NULL, " &_
	  " 	Descr_Bussola_spa ntext NULL ; " &_
	  " DROP VIEW dbo.viewEventi ; " &_
	  " CREATE VIEW dbo.viewEventi AS " &_
	  " 	SELECT ID, id_categoria, id_ev_spec, ev_tel_org, ingresso_intero, ingresso_ridotto, " &_
	  "			titolo, titol_eng, titol_fra, titol_spa, titol_ted, " &_
	  "			descrizione, descr_eng, descr_fra, descr_ted, descr_spa, " &_
	  "			orario, orario_eng, orario_fra, orario_ted, orario_spa, " &_
	  "			Info, info_eng, info_fra, info_ted, info_spa, " &_
	  "			descr_luogo, descr_luogo_eng, descr_luogo_fra, descr_luogo_ted, descr_luogo_spa " &_
	  "		FROM Eventi; "
CALL DB.Execute(sql, 80)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 81
'...........................................................................................
'aggiunge campi per la gestione delle mappe su vista locali e servizi
'...........................................................................................
sql = " DROP VIEW dbo.viewLocalieServizi; " &_
	  " CREATE VIEW dbo.viewLocalieServizi AS " &_
	  "		SELECT dbo.LocalieServizi.*, dbo.Tipi_LS.tipo_nome_it, dbo.Tipi_LS.tipo_nome_eng " &_
	  "			FROM dbo.LocalieServizi INNER JOIN dbo.Tipi_LS ON dbo.LocalieServizi.Tipo = dbo.Tipi_LS.id_tipoutil "
CALL DB.Execute(sql, 81)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 82
'...........................................................................................
'aggiunge campi per la gestione delle mappe su vista per notizie utili
'...........................................................................................
sql = " DROP VIEW dbo.viewAllNotUtili; " &_
	  " CREATE VIEW dbo.viewAllNotUtili AS " &_
	  "		SELECT * FROM dbo.Not_util INNER JOIN dbo.Tipi_notutil ON dbo.Not_util.Tipo = dbo.Tipi_notutil.id_tipoutil "
CALL DB.Execute(sql, 82)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 83
'...........................................................................................
'aggiunge campi per la gestione delle mappe su tabella locali e servizi
'...........................................................................................
sql = " ALTER TABLE LocalieServizi ADD linkmappe nvarchar(6) NULL"
CALL DB.Execute(sql, 83)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 84
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom
'...........................................................................................
sql = " ALTER TABLE tb_Indirizzario ADD " & _
	  " 	SyncroKey nvarchar(50) NULL, " & _
	  " 	SyncroTable nvarchar(50) NULL, " & _
	  " 	SyncroApplication INT NULL"
CALL DB.Execute(sql, 84)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 85
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom
'...........................................................................................
sql = " ALTER TABLE tb_ValoriNumeri ADD SyncroField nvarchar(50) NULL"
CALL DB.Execute(sql, 85)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 86
'...........................................................................................
'aggiunge campi in lingua per notizie utili
'...........................................................................................
sql = " ALTER TABLE Not_Util ADD " &_
	  "		Descr_fra ntext NULL, " &_
	  "		Descr_ted ntext NULL, " &_
	  "		Descr_spa ntext NULL"
CALL DB.Execute(sql, 86)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 87
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_rubriche ADD " & _
	  " SyncroTable nvarchar(50) NULL, " &_
	  " SyncroFilter INTEGER NULL"
CALL DB.Execute(sql, 87)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 88
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_indirizzario ADD LocalitaElencoIndirizzi nvarchar(100) NULL"
CALL DB.Execute(sql, 88)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 89
'...........................................................................................
'Aggiunge gruppi di lavoro al NextCom e svuotamento dati presenti
'...........................................................................................
sql = " DELETE FROM tb_gruppi; " &_
	  " DELETE FROM tb_rubriche; " &_
	  " DELETE FROM tb_indirizzario; " &_
	  " DELETE FROM tb_email; " &_
	  "	SET IDENTITY_INSERT tb_gruppi ON; " &_
	  " INSERT tb_gruppi(id_gruppo, nome_gruppo) VALUES(5, 'Segreteria'); " &_
	  " INSERT tb_gruppi(id_gruppo, nome_gruppo) VALUES(6, 'Ufficio statistiche'); " &_
	  " INSERT tb_gruppi(id_gruppo, nome_gruppo) VALUES(7, 'Ufficio Relazioni con il Pubblico'); " &_
	  " INSERT tb_gruppi(id_gruppo, nome_gruppo) VALUES(8, 'Ufficio Personale'); " &_
	  " INSERT tb_gruppi(id_gruppo, nome_gruppo) VALUES(9, 'Redazione') " &_
	  "	SET IDENTITY_INSERT tb_gruppi OFF "
CALL DB.Execute(sql, 89)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 90
'...........................................................................................
'Aggiunge rubriche di sistema per sincronizzazione con dati del portale
'...........................................................................................
sql = " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Strutture ricettive', 'Stru_ric', NULL, 1, 1) " &_
	  " INSERT INTO tb_rel_GruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Spiagge', 'Spiagge', NULL, 1, 1) " &_
	  " INSERT INTO tb_rel_GruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Luoghi', 'Luoghi', NULL, 1, 1) " &_
	  " INSERT INTO tb_rel_GruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Locali & Servizi', 'LocaliEServizi', NULL, 1, 1) " &_
	  " INSERT INTO tb_rel_GruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Notizie utili', 'Not_Util', NULL, 1, 1) " &_
	  " INSERT INTO tb_rel_GruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; "
CALL DB.Execute(sql, 90)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 91
'...........................................................................................
'modifica stored procedure di cancellazione per sincronizzazione con next-Com
'...........................................................................................
sql = " ALTER PROCEDURE dbo.DELETE_Stru_Ric (@ID_Albergo int) AS " & vbcRLF &_
	  "		DELETE FROM RelazPS WHERE ID_StruRic = @ID_Albergo " & vbcRLF &_
	  "		DELETE FROM Rel_ric_caratt WHERE rel_id_struric= @ID_Albergo " & vbcRLF &_
	  "		DELETE FROM Rel_Ric_Servizi WHERE rel_ServRic_idRic = @ID_Albergo " & vbcRLF &_
	  "		DELETE FROM tb_Indirizzario WHERE SyncroTable LIKE 'Stru_ric' AND SyncroKey LIKE CAST(@ID_Albergo AS nvarchar(50)) " & vbcRLF &_
	  "		DELETE FROM Stru_ric WHERE ID_Albergo = @ID_Albergo; " &_
	  " ALTER PROCEDURE dbo.DELETE_Spiagge (@ID_Spiagge int) AS " & vbcrlf & _
	  "		DELETE FROM tb_Indirizzario WHERE SyncroTable LIKE 'Spiagge' AND SyncroKey LIKE CAST(@ID_Spiagge AS nvarchar(50)) " & vbcRLF &_
	  "		DELETE FROM Spiagge WHERE ID_Spiagge=@ID_Spiagge; " & _
	  "	ALTER PROCEDURE dbo.DELETE_Not_Util ( @Not_ID int ) AS " & vbCrLf & _
	  "		DELETE FROM Rel_sottotipi_NotUtil WHERE rel_sTipoNot_Not= @Not_ID " & VbCrLF &_
	  "		DELETE FROM tb_Indirizzario WHERE SyncroTable LIKE 'Not_Util' AND SyncroKey LIKE CAST(@Not_ID AS nvarchar(50)) " & vbcRLF &_
	  " 	DELETE FROM Not_Util WHERE id_Util = @Not_ID; " & _ 
	  " ALTER PROCEDURE dbo.DELETE_Luoghi ( @ID int) AS " & VbCrLf & _
	  "		DELETE FROM Imag_lu WHERE id_Luogo = @ID " & VbCrLf & _
	  "		DELETE FROM doveAccade WHERE id_luogo = @ID " & VbCrLf & _
	  "		DELETE FROM tb_Indirizzario WHERE SyncroTable LIKE 'Luoghi' AND SyncroKey LIKE CAST(@ID AS nvarchar(50)) " & vbcRLF &_
	  "		DELETE FROM Luoghi WHERE ID = @ID; " & _
	  "	ALTER PROCEDURE dbo.DELETE_LS (@LS_ID int) AS " & vbCrLf &_
	  "		DELETE FROM Rel_sottotipi_LS WHERE rel_sLS_LocaleServizio= @LS_ID " & vbCrLf & _
	  "		DELETE FROM tb_Indirizzario WHERE SyncroTable LIKE 'LocaliEServizi' AND SyncroKey LIKE CAST(@LS_ID AS nvarchar(50)) " & vbcRLF &_
	  "		DELETE FROM LocalieServizi WHERE id_LS= @LS_ID "
CALL DB.Execute(sql, 91)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 92
'...........................................................................................
'inserisce le rubriche collegate con le tipologie di luoghi, notizie utili, locali e servizi e ricettivita'
'...........................................................................................
sql = " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Strutture ricettive - ' + Denominazione_it, 'Stru_ric', IDTipo, 1, 1 FROM Tipi_Ric; " & _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Locali & Servizi - ' + tipo_nome_it, 'LocaliEServizi', id_tipoUtil, 1, 1 FROM Tipi_LS; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Notizie utili - ' + tipo_nome_it, 'Not_Util', id_tipoUtil, 1, 1 FROM Tipi_NotUtil; " & _
  	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Luoghi - ' + tipo_luogo, 'Luoghi', IDL, 1, 1 FROM TipoLuoghi WHERE Visibile=1; " & _
	  " DELETE FROM tb_rel_gruppirubriche; " & _
	  " INSERT INTO tb_rel_gruppirubriche (id_dellarubrica, id_gruppo_assegnato) " &_
	  " 	SELECT tb_rubriche.id_rubrica, tb_gruppi.id_gruppo FROM tb_rubriche, tb_gruppi "
CALL DB.Execute(sql, 92)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 93
'...........................................................................................
'crea i trigger che inseriscono le rubriche collegate con i tipi (SOLO PER INSERIMENTO)
'...........................................................................................
sql = " CREATE TRIGGER dbo.T_TipoLuoghi_INSERT ON TipoLuoghi AFTER INSERT AS " & vbCrLf &_
	  "		DECLARE @filter bit " & vbCrLf &_
	  "		SELECT @filter=visibile from INSERTED " & vbCrLf &_
	  "		if (@filter=1) BEGIN " & vbCrLf &_
	  "			INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "				SELECT 'Luoghi - ' + tipo_luogo, 'Luoghi', IDL, 1, 1 FROM INSERTED " & vbCrLf &_
	  "			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "				SELECT @@IDENTITY, id_gruppo FROM tb_gruppi " & vbCrLf &_
	  "		END ;" & vbCrLf &_
	  "	CREATE TRIGGER dbo.T_Tipi_NotUtil_INSERT ON Tipi_NotUtil AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Notizie utili - ' + tipo_nome_it, 'Not_Util', id_tipoutil, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi ; " & vbCrLf &_
	  "	CREATE TRIGGER dbo.T_Tipi_LS_INSERT ON Tipi_LS AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Locali & Servizi - ' + tipo_nome_it, 'LocalieServizi', id_tipoutil, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi ; " & vbCrLf &_
	  "	CREATE TRIGGER dbo.T_Tipi_ric_INSERT ON Tipi_ric AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Str. ric. - ' + Denominazione_it, 'Stru_ric', IDTipo, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi"
CALL DB.Execute(sql, 93)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 94
'...........................................................................................
'modifica delle procedure che cancellano i tipi collegati alle rubriche
'...........................................................................................
sql = "ALTER PROCEDURE dbo.DELETE_Tipi_ric (@IDTipo int) AS " & vbcRLF &_
	  "		DELETE FROM RelazPS WHERE ID_StruRic IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Rel_ric_caratt WHERE rel_id_struric IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Rel_Ric_Servizi WHERE rel_ServRic_idRic IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Stru_ric WHERE Tipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Rel_Serv_Tipiric WHERE rel_tRicServ_idTipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Rel_TipiStr_Caratt WHERE rel_id_Tipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroTable LIKE 'Stru_ric' AND SyncroFilter = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Tipi_Ric WHERE IDTipo = @IDTipo ; " & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_Tipo_LS (@Tip_ID int) AS " & vbcRLF &_
	  "		DELETE FROM rel_sottoTipi_LS WHERE rel_sLS_LocaleServizio IN (SELECT id_LS FROM LocalieServizi WHERE Tipo = @Tip_ID) " & vbcRLF &_
	  "		DELETE FROM LocalieServizi WHERE Tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM SottoTipi_LocServ WHERE ref_tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroTable LIKE 'LocalieServizi' AND SyncroFilter = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM Tipi_LS WHERE id_TipoUtil = @Tip_ID ;" & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_Tipo_Not_Util (@Tip_ID int) AS " & vbcRLF &_
	  "		DELETE FROM rel_sottoTipi_NotUtil WHERE rel_sTipoNot_Not IN (SELECT id_Util FROM Not_Util WHERE Tipo = @Tip_ID) " & vbcRLF &_
	  "		DELETE FROM Not_Util WHERE Tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM SottoTipi_NotUtil WHERE ref_tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroTable LIKE 'Not_Util' AND SyncroFilter = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM Tipi_NotUtil WHERE id_TipoUtil = @Tip_ID ;" & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_TipoLuoghi ( @IDL int) AS " & vbcRLF &_
	  "		DELETE FROM Imag_lu WHERE id_Luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM doveAccade WHERE id_luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM Luoghi WHERE ID IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroTable LIKE 'Luoghi' AND SyncroFilter = @IDL " & vbcRLF &_
	  "		DELETE FROM TipoLuoghi WHERE IDL = @IDL ;"
CALL DB.Execute(sql, 94)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 95
'...........................................................................................
'crea i trigger che aggiornano le rubriche collegate con i tipi (SOLO PER MODIFICA)
'...........................................................................................
sql = " CREATE TRIGGER dbo.T_Tipi_NotUtil_UPDATE ON Tipi_NotUtil AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Tipo_nome_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_tipoutil, @nome=tipo_nome_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Notizie utili - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroTable LIKE 'Not_Util' AND SyncroFilter=@ID " & vbcRLF & _
	  "		END; " & _
	  " CREATE TRIGGER dbo.T_Tipi_LS_UPDATE ON Tipi_LS AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Tipo_nome_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_tipoutil, @nome=tipo_nome_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Locali & Servizi - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroTable LIKE 'LocalieServizi' AND SyncroFilter=@ID " & vbcRLF & _
	  "		END; " & _
	  " CREATE TRIGGER dbo.T_Tipi_ric_UPDATE ON Tipi_ric AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Denominazione_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=IDTipo, @nome=Denominazione_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Str. ric. - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroTable LIKE 'Stru_ric' AND SyncroFilter=@ID " & vbcRLF & _
	  "		END; " & vbcRLF & _
	  " CREATE TRIGGER dbo.T_TipoLuoghi_UPDATE ON TipoLuoghi AFTER UPDATE AS " & vbcRLF & _
	  "		DECLARE @ID int " & vbcRLF & _
	  "		DECLARE @nome nvarchar(255) " & vbcRLF & _
	  "		DECLARE @visibile bit " & vbcRLF & _
	  "		SELECT @ID=IDL, @nome=Tipo_luogo, @Visibile=visibile FROM INSERTED " & vbcRLF & _
	  "		if (UPDATE(visibile)) BEGIN " & vbcRLF & _
	  "			if (@visibile=1) BEGIN " & vbcRLF & _
	  "				DECLARE @COUNT int " & vbcRLF & _
	  "				SELECT @COUNT=COUNT(*) FROM tb_rubriche WHERE SyncroTable LIKE 'Luoghi' AND SyncroFilter=@ID " & vbcRLF & _
	  "				IF (@COUNT>0) BEGIN " & vbcRLF & _
	  "					UPDATE tb_rubriche SET nome_rubrica='Luoghi - ' + @Nome  " & vbcRLF & _
	  "						WHERE SyncroTable LIKE 'Luoghi' AND SyncroFilter=@ID " & vbcRLF & _
	  "				END " & vbcRLF & _
	  "				ELSE BEGIN " & vbcRLF & _
	  "					DECLARE @ID_RUBRICA int " & vbcRLF & _
	  "					INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilter, locked_rubrica, rubrica_esterna)  " & vbcRLF & _
	  "						VALUES ('Luoghi - ' + @Nome, 'Luoghi', @ID, 1, 1) " & vbcRLF & _
	  "					SET @ID_RUBRICA = @@IDENTITY " & vbcRLF & _
	  "		 			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbcRLF & _
	  "						SELECT @ID_RUBRICA, id_gruppo FROM tb_gruppi " & vbcRLF & _
	  "					INSERT INTO rel_rub_ind (id_rubrica, id_indirizzo) " & vbcRLF & _
	  "						SELECT @ID_RUBRICA, IDElencoIndirizzi FROM tb_indirizzario  " & vbcRLF & _
	  "							WHERE SyncroTable LIKE 'Luoghi' " & vbcRLF & _
	  "							AND CAST(SyncroKey AS int) IN (SELECT ID FROM luoghi WHERE id_tipo=8) " & vbcRLF & _
	  "				END " & vbcRLF & _
	  "			END " & vbcRLF & _
	  "			ELSE BEGIN " & vbcRLF & _
	  "				DELETE FROM tb_rubriche WHERE SyncroTable LIKE 'Luoghi' AND SyncroFilter=@ID " & vbcRLF & _
	  "			END " & vbcRLF & _
	  "		END; "
CALL DB.Execute(sql, 95)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 96
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " ALTER TABLE tb_rubriche ADD " & _
	  " SyncroFilterTable nvarchar(50) NULL, " &_
	  " SyncroFilterKey INTEGER NULL; " &_
	  " UPDATE tb_rubriche SET SyncroFilterKey = SyncroFilter, " &_
	  " 	SyncroFilterTable = CASE WHEN SyncroTable LIKE 'Not_Util' THEN 'Tipi_notUtil' " &_
	  " 	 	 WHEN SyncroTable LIKE 'LocalieServizi' THEN 'Tipi_LS' " &_
	  " 		 WHEN SyncroTable LIKE 'Stru_ric' THEN 'Tipi_ric' " &_
	  " 		 WHEN SyncroTable LIKE 'Luoghi' THEN 'TipoLuoghi' END;"
CALL DB.Execute(sql, 96)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 97
'...........................................................................................
'aggiunge campi per la sincronizzazione dati con applicativi esterni al NextCom su tb_rubriche
'...........................................................................................
sql = " UPDATE tb_rubriche SET SyncroFilterTable=NULL WHERE SyncroFilter IS NULL; " &_
	  " ALTER TABLE tb_rubriche DROP COLUMN SyncroFilter; "
CALL DB.Execute(sql, 97)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 98
'...........................................................................................
'modifica i trigger per variazione campi di sincronizzazione
'...........................................................................................
sql = " ALTER TRIGGER T_TipoLuoghi_INSERT ON TipoLuoghi AFTER INSERT AS " & vbCrLf &_
	  "		DECLARE @filter bit " & vbCrLf &_
	  "		SELECT @filter=visibile from INSERTED " & vbCrLf &_
	  "		if (@filter=1) BEGIN " & vbCrLf &_
	  "			INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "				SELECT 'Luoghi - ' + tipo_luogo, 'Luoghi', 'TipoLuoghi', IDL, 1, 1 FROM INSERTED " & vbCrLf &_
	  "			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "				SELECT @@IDENTITY, id_gruppo FROM tb_gruppi " & vbCrLf &_
	  "		END ;" & vbCrLf &_
	  "	ALTER TRIGGER T_Tipi_NotUtil_INSERT ON Tipi_NotUtil AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Notizie utili - ' + tipo_nome_it, 'Not_Util', 'Tipi_NotUtil', id_tipoutil, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi ; " & vbCrLf &_
	  "	ALTER TRIGGER T_Tipi_LS_INSERT ON Tipi_LS AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Locali & Servizi - ' + tipo_nome_it, 'LocalieServizi', 'Tipi_LS', id_tipoutil, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi ; " & vbCrLf &_
	  "	ALTER TRIGGER T_Tipi_ric_INSERT ON Tipi_ric AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "			SELECT 'Str. ric. - ' + Denominazione_it, 'Stru_ric', 'Tipi_ric', IDTipo, 1, 1 FROM INSERTED " & vbCrLf &_
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi"
CALL DB.Execute(sql, 98)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 99
'...........................................................................................
'modifica delle procedure che cancellano i tipi collegati alle rubriche
'...........................................................................................
sql = "ALTER PROCEDURE dbo.DELETE_Tipi_ric (@IDTipo int) AS " & vbcRLF &_
	  "		DELETE FROM RelazPS WHERE ID_StruRic IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Rel_ric_caratt WHERE rel_id_struric IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Rel_Ric_Servizi WHERE rel_ServRic_idRic IN (SELECT ID_Albergo FROM Stru_Ric WHERE Tipo= @IDTipo) " & vbcRLF &_
	  "		DELETE FROM Stru_ric WHERE Tipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Rel_Serv_Tipiric WHERE rel_tRicServ_idTipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Rel_TipiStr_Caratt WHERE rel_id_Tipo = @IDTipo " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'Tipi_Ric' AND SyncroFilterKey = @IDTipo " & vbcRLF &_
	  "		DELETE FROM Tipi_Ric WHERE IDTipo = @IDTipo ; " & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_Tipo_LS (@Tip_ID int) AS " & vbcRLF &_
	  "		DELETE FROM rel_sottoTipi_LS WHERE rel_sLS_LocaleServizio IN (SELECT id_LS FROM LocalieServizi WHERE Tipo = @Tip_ID) " & vbcRLF &_
	  "		DELETE FROM LocalieServizi WHERE Tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM SottoTipi_LocServ WHERE ref_tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'Tipi_LS' AND SyncroFilterKey = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM Tipi_LS WHERE id_TipoUtil = @Tip_ID ;" & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_Tipo_Not_Util (@Tip_ID int) AS " & vbcRLF &_
	  "		DELETE FROM rel_sottoTipi_NotUtil WHERE rel_sTipoNot_Not IN (SELECT id_Util FROM Not_Util WHERE Tipo = @Tip_ID) " & vbcRLF &_
	  "		DELETE FROM Not_Util WHERE Tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM SottoTipi_NotUtil WHERE ref_tipo = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'Tipi_NotUtil' AND SyncroFilterKey = @Tip_ID " & vbcRLF &_
	  "		DELETE FROM Tipi_NotUtil WHERE id_TipoUtil = @Tip_ID ;" & vbcRLF &_
	  "	ALTER PROCEDURE dbo.DELETE_TipoLuoghi ( @IDL int) AS " & vbcRLF &_
	  "		DELETE FROM Imag_lu WHERE id_Luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM doveAccade WHERE id_luogo IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM Luoghi WHERE ID IN (SELECT ID FROM Luoghi WHERE id_tipo=@IDL) " & vbcRLF &_
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'TipoLuoghi' AND SyncroFilterKey = @IDL " & vbcRLF &_
	  "		DELETE FROM TipoLuoghi WHERE IDL = @IDL ;"
CALL DB.Execute(sql, 99)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 100
'...........................................................................................
'modifica i trigger che aggiornano le rubriche collegate con i tipi (SOLO PER MODIFICA)
'...........................................................................................
sql = " ALTER TRIGGER T_Tipi_NotUtil_UPDATE ON Tipi_NotUtil AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Tipo_nome_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_tipoutil, @nome=tipo_nome_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Notizie utili - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroFilterTable LIKE 'Tipi_NotUtil' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "		END; " & _
	  " ALTER TRIGGER T_Tipi_LS_UPDATE ON Tipi_LS AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Tipo_nome_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_tipoutil, @nome=tipo_nome_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Locali & Servizi - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroFilterTable LIKE 'Tipi_LS' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "		END; " & _
	  " ALTER TRIGGER T_Tipi_ric_UPDATE ON Tipi_ric AFTER UPDATE AS " & vbcRLF & _
	  "		if (UPDATE(Denominazione_it)) BEGIN" & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=IDTipo, @nome=Denominazione_it FROM INSERTED " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica='Str. ric. - ' + @Nome " & vbcRLF & _
	  "				WHERE SyncroFilterTable LIKE 'Tipi_ric' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "		END; " & vbcRLF & _
	  " ALTER TRIGGER T_TipoLuoghi_UPDATE ON TipoLuoghi AFTER UPDATE AS " & vbcRLF & _
	  "		DECLARE @ID int " & vbcRLF & _
	  "		DECLARE @nome nvarchar(255) " & vbcRLF & _
	  "		DECLARE @visibile bit " & vbcRLF & _
	  "		SELECT @ID=IDL, @nome=Tipo_luogo, @Visibile=visibile FROM INSERTED " & vbcRLF & _
	  "		if (UPDATE(visibile)) BEGIN " & vbcRLF & _
	  "			if (@visibile=1) BEGIN " & vbcRLF & _
	  "				DECLARE @COUNT int " & vbcRLF & _
	  "				SELECT @COUNT=COUNT(*) FROM tb_rubriche WHERE SyncroFilterTable LIKE 'TipoLuoghi' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "				IF (@COUNT>0) BEGIN " & vbcRLF & _
	  "					UPDATE tb_rubriche SET nome_rubrica='Luoghi - ' + @Nome  " & vbcRLF & _
	  "						WHERE SyncroFilterTable LIKE 'TipoLuoghi' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "				END " & vbcRLF & _
	  "				ELSE BEGIN " & vbcRLF & _
	  "					DECLARE @ID_RUBRICA int " & vbcRLF & _
	  "					INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna)  " & vbcRLF & _
	  "						VALUES ('Luoghi - ' + @Nome, 'Luoghi', 'TipoLuoghi', @ID, 1, 1) " & vbcRLF & _
	  "					SET @ID_RUBRICA = @@IDENTITY " & vbcRLF & _
	  "		 			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbcRLF & _
	  "						SELECT @ID_RUBRICA, id_gruppo FROM tb_gruppi " & vbcRLF & _
	  "					INSERT INTO rel_rub_ind (id_rubrica, id_indirizzo) " & vbcRLF & _
	  "						SELECT @ID_RUBRICA, IDElencoIndirizzi FROM tb_indirizzario  " & vbcRLF & _
	  "							WHERE SyncroTable LIKE 'Luoghi' " & vbcRLF & _
	  "							AND CAST(SyncroKey AS int) IN (SELECT ID FROM luoghi WHERE id_tipo=8) " & vbcRLF & _
	  "				END " & vbcRLF & _
	  "			END " & vbcRLF & _
	  "			ELSE BEGIN " & vbcRLF & _
	  "				DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'TipoLuoghi' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "			END " & vbcRLF & _
	  "		END; "
CALL DB.Execute(sql, 100)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 101
'...........................................................................................
'aggiunge trigger e modifica stored procedure per sincronizzazione rubriche <-> sottotipi notizie utili 
' e rubriche <-> sottotipi locali e servizi
'...........................................................................................
sql = " CREATE TRIGGER dbo.T_SottoTipi_NotUtil_INSERT ON SottoTipi_NotUtil AFTER INSERT AS " & vbcRLF & _
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbcRLF & _
	  "			SELECT 'Notizie utili - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'Not_Util', 'SottoTipi_NotUtil', id_sottipo, 1, 1 " & vbcRLF & _
	  "				FROM INSERTED INNER JOIN Tipi_notutil ON INSERTED.ref_tipo = Tipi_NotUtil.id_tipoUtil " & vbcRLF & _
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbcRLF & _
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi;  " & vbcRLF & _
	  " CREATE TRIGGER dbo.T_SottoTipi_NotUtil_UPDATE ON SottoTipi_NotUtil AFTER UPDATE AS  " & vbcRLF & _
	  "		IF (UPDATE(sottip_nome_it)) BEGIN " & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_sottipo, @nome='Notizie utili - ' + tipo_nome_it + ' - ' + sottip_nome_it " & vbcRLF & _
	  "				FROM INSERTED INNER JOIN Tipi_notutil ON INSERTED.ref_tipo = Tipi_NotUtil.id_tipoUtil " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica=@Nome " & vbcRLF & _
	  "				WHERE SyncroFilterTable LIKE 'SottoTipi_NotUtil' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "		END; " & vbcRLF & _
	  " ALTER PROCEDURE dbo.Delete_SUbTipo_Not( @Tip_ID int )AS " & vbcRLF & _
	  "		DELETE FROM rel_sottoTipi_NotUtil WHERE rel_sTipoNot_sTipo = @Tip_ID " & vbcRLF & _
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'SottoTipi_NotUtil' AND SyncroFilterKey = @Tip_ID " & vbcRLF & _
	  "		DELETE FROM SottoTipi_NotUtil WHERE id_sottipo = @Tip_ID; " & vbcRLF & _
	  " CREATE TRIGGER dbo.T_SottoTipi_LocServ_INSERT ON SottoTipi_LocServ AFTER INSERT AS " & vbcRLF & _
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbcRLF & _
	  "			SELECT 'Locali & Servizi - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'LocaliEServizi', 'SottoTipi_LocServ', id_sottipo, 1, 1 " & vbcRLF & _
	  "				FROM INSERTED INNER JOIN Tipi_LS ON INSERTED.ref_tipo = Tipi_LS.id_tipoUtil " & vbcRLF & _
	  "		INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbcRLF & _
	  "			SELECT @@IDENTITY, id_gruppo FROM tb_gruppi;  " & vbcRLF & _
	  " CREATE TRIGGER dbo.T_SottoTipi_LocServ_UPDATE ON SottoTipi_LocServ AFTER UPDATE AS  " & vbcRLF & _
	  "		IF (UPDATE(sottip_nome_it)) BEGIN " & vbcRLF & _
	  "			DECLARE @ID int " & vbcRLF & _
	  "			DECLARE @nome nvarchar(250) " & vbcRLF & _
	  "			SELECT @ID=id_sottipo, @nome='Locali & Servizi - ' + tipo_nome_it + ' - ' + sottip_nome_it " & vbcRLF & _
	  "				FROM INSERTED INNER JOIN Tipi_LS ON INSERTED.ref_tipo = Tipi_LS.id_tipoUtil " & vbcRLF & _
	  "			UPDATE tb_rubriche SET nome_rubrica=@Nome " & vbcRLF & _
	  "				WHERE SyncroFilterTable LIKE 'SottoTipi_LocServ' AND SyncroFilterKey=@ID " & vbcRLF & _
	  "		END; " & vbcRLF & _
	  " ALTER PROCEDURE dbo.Delete_SubTipo_LS( @Tip_ID int )AS " & vbcRLF & _
	  "		DELETE FROM rel_sottoTipi_LS WHERE rel_sLS_sTipo = @Tip_ID " & vbcRLF & _
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'SottoTipi_LocServ' AND SyncroFilterKey = @Tip_ID " & vbcRLF & _
	  "		DELETE FROM SottoTipi_LocServ WHERE id_sottipo = @Tip_ID"
CALL DB.Execute(sql, 101)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 102
'...........................................................................................
'aggiunge trigger per sincronizzazione sottoitipi-notizie <-> rubriche-contatti sincronizzati
'...........................................................................................
sql = " CREATE TRIGGER T_rel_sottotipi_NotUtil_INSERT ON rel_sottotipi_NotUtil AFTER INSERT AS" & vbcRLF & _
	  "		DECLARE @ID_INDIRIZZO int" & vbcRLF & _
	  "		DECLARE @ID_NOT int" & vbcRLF & _
	  "		DECLARE @ID_SOTTOTIPO int" & vbcRLF & _
	  "		SELECT @ID_NOT=rel_sTipoNot_Not, @ID_SOTTOTIPO=rel_sTipoNot_sTipo FROM INSERTED" & vbcRLF & _
	  "		SELECT @ID_INDIRIZZO = IDElencoIndirizzi FROM tb_indirizzario " & vbcRLF & _
	  "			WHERE SyncroTable LIKE 'Not_Util' AND CAST(SyncroKey AS int)=@ID_NOT" & vbcRLF & _
	  "		INSERT INTO rel_rub_ind(id_indirizzo, id_rubrica)" & vbcRLF & _
	  "			SELECT @ID_INDIRIZZO, id_Rubrica FROM tb_rubriche " & vbcRLF & _
	  "				WHERE SyncroTable LIKE 'Not_Util' AND SyncroFilterTable='SottoTipi_NotUtil' AND " & vbcRLF & _
	  "				SyncroFilterKey IN (@ID_SOTTOTIPO) ;" & vbcRLF & _
	  "	CREATE TRIGGER T_rel_sottotipi_NotUtil_DELETE ON rel_sottotipi_NotUtil AFTER DELETE AS" & vbcRLF & _
	  "		DELETE FROM rel_rub_ind WHERE " & vbcRLF & _
	  "			id_indirizzo IN (SELECT IDElencoIndirizzi FROM tb_indirizzario INNER JOIN DELETED ON " & vbcRLF & _
	  "					CAST(tb_indirizzario.SyncroKey AS int) = DELETED.rel_sTipoNot_not)" & vbcRLF & _
	  "			AND id_rubrica IN (SELECT id_rubrica FROM tb_rubriche INNER JOIN DELETED ON" & vbcRLF & _
	  "					tb_rubriche.SyncroFilterKey = DELETED.rel_sTipoNot_sTipo); "
CALL DB.Execute(sql, 102)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 103
'...........................................................................................
'inserisce rubriche collegate a tutti i sottotipi di locali e servizi e notizie utili
'...........................................................................................
sql = " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " &_
	  "		SELECT 'Notizie utili - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'Not_Util', 'SottoTipi_NotUtil', id_sottipo, 1,1 " &_
	  "			FROM SottoTipi_NotUtil INNER JOIN Tipi_notUtil ON SottoTipi_NotUtil.ref_tipo=tipi_NotUtil.id_tipoUtil ;" &_
	  "	INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  "		SELECT 'Locali & Servizi - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'LocaliEServizi', 'SottoTipi_LocServ', id_sottipo, 1,1 " &_
	  "			FROM SottoTipi_LocServ INNER JOIN Tipi_LS ON SottoTipi_LocServ.ref_tipo=Tipi_LS.id_tipoUtil ;" &_
	  "	INSERT INTO tb_rel_gruppirubriche (id_dellarubrica, id_gruppo_assegnato) " &_
	  " 	SELECT tb_rubriche.id_rubrica, tb_gruppi.id_gruppo FROM tb_rubriche, tb_gruppi " &_
	  "			WHERE tb_rubriche.SyncroFilterTable LIKE 'SottoTipi_NotUtil' OR tb_rubriche.SyncroFilterTable LIKE 'SottoTipi_LocServ'"
CALL DB.Execute(sql, 103)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 104
'...........................................................................................
'ripulisce tabelle dotazioni e servizi delle strutture aggiornate automaticamente dal WS
'...........................................................................................
sql = " DELETE FROM rel_ric_caratt WHERE rel_id_struRic IN " &_
	  "		(SELECT ID_Albergo FROM Stru_ric INNER JOIN Tipi_ric ON Stru_ric.Tipo = Tipi_ric.IDTipo WHERE ID_Tipi_Provincia<>''); " &_
	  "	DELETE FROM rel_ric_Servizi WHERE rel_ServRic_idRic IN " &_
	  "		(SELECT ID_Albergo FROM Stru_ric INNER JOIN Tipi_ric ON Stru_ric.Tipo = Tipi_ric.IDTipo WHERE ID_Tipi_Provincia<>''); "
CALL DB.Execute(sql, 104)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 105
'...........................................................................................
'sposta id rubriche esistenti oltre 1000 per impedire incrocio dati con import
'...........................................................................................
sql = " SET IDENTITY_INSERT tb_rubriche ON " &_
	  " INSERT INTO tb_rubriche(id_Rubrica, nome_Rubrica, note_Rubrica, locked_rubrica, rubrica_esterna, SyncroTable, SyncroFilterTable, SyncroFilterKey) " &_
	  "		SELECT (id_rubrica + 1000), nome_rubrica, note_rubrica, locked_rubrica, rubrica_esterna, SyncroTable, SyncroFilterTable, SyncroFilterKey FROM tb_rubriche " &_
	  "	UPDATE rel_rub_ind SET id_rubrica = (id_rubrica + 1000) " &_
	  "	UPDATE tb_rel_gruppirubriche SET id_dellaRubrica = (id_dellaRubrica + 1000) " &_
	  " DELETE tb_rubriche WHERE id_rubrica < 1000 " &_
	  " SET IDENTITY_INSERT tb_rubriche OFF "
CALL DB.Execute(sql, 105)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 106
'...........................................................................................
'sposta id rubriche esistenti oltre 1000 per impedire incrocio dati con import
'...........................................................................................
sql = " ALTER TABLE tb_ValoriNumeri ALTER COLUMN ValoreNumero nvarchar(250) "
CALL DB.Execute(sql, 106)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 107
'...........................................................................................
'importa dati da database indirizzario (richiede dbindirizzario.mdb in cartella database
'...........................................................................................
'AGGIORNAMENTO rimosso dal framework per rimozione applicativo AptAdmin il 03/12/2007
'<!--#include file="subscripts/Update_govenice_107_Inport_Indirizzario.asp"-->
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 107)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 108
'...........................................................................................
'sposta id rubriche esistenti oltre 1000 per impedire incrocio dati con import
'...........................................................................................
sql = " CREATE  CLUSTERED  INDEX IX_rel_rub_ind ON dbo.rel_rub_ind(id_indirizzo); " &_
	  " CREATE  CLUSTERED  INDEX IX_tb_ValoriNumeri ON dbo.tb_ValoriNumeri(id_Indirizzario)" 
CALL DB.Execute(sql, 108)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 109
'...........................................................................................
'completa sincronizzazione record vari con contati (LUOGHI, strutture Ricettive, Notizie utili)
'...........................................................................................
'AGGIORNAMENTO rimosso dal framework per rimozione applicativo AptAdmin il 03/12/2007
'<!--#include file="subscripts/Update_govenice_109_Completa_syncro.asp"-->
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 109)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 110
'...........................................................................................
'pulisce valori numeri non validi
'...........................................................................................
sql = " DELETE FROM tb_ValoriNumeri WHERE ISNULL(ValoreNumero, '')=''; "	  
CALL DB.Execute(sql, 110)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 111
'...........................................................................................
'esegue modifiche alle rubriche
'...........................................................................................
sql = " UPDATE tb_rubriche SET nome_rubrica='APT - Provincia di Venezia' WHERE id_rubrica=23; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='APT - Fuuori Provincia di Venezia' WHERE id_rubrica=269; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Aziende & Societ' + NCHAR(224) + ' varie' WHERE id_rubrica=9; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Autorit' + NCHAR(224) + ' ambito APT' WHERE id_rubrica=263; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Agenzie di viaggio & Tour Operator', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Cultura Italia - Biblioteche', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Musei' WHERE id_rubrica=24; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Musica' WHERE id_rubrica=26; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Enti & Societ' + NCHAR(224) WHERE id_rubrica=34; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Estero' WHERE id_rubrica=7; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=34 WHERE id_rubrica=27; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 27; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Venezia / Mestre - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Venezia / Mestre - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Riviera del Brenta - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Riviera del Brenta - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Terra dei Tiepolo - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Terra dei Tiepolo - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Cavallino - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Cavallino - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Quarto d''Altino - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Quarto d''Altino - Assessori Turismo & Cultura', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Assessori turismo ambito APT' WHERE id_rubrica=5; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Sindaci' WHERE id_rubrica=4; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Assessori Turismo fuori APT' WHERE id_rubrica=268; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consigli di Quartiere & Municipalit' + NCHAR(224) WHERE id_rubrica=17; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consolati & Ambasciate Italiane all''estero' WHERE id_rubrica=12; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consolati & Ambasciate Estere in Italia' WHERE id_rubrica=11; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Locale' WHERE id_rubrica=3; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=3 WHERE id_rubrica=13; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 13; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Nazionale' WHERE id_rubrica=264; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=264 WHERE id_rubrica=28; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=264 WHERE id_rubrica=29; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 28; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 29; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Estera' WHERE id_rubrica=265; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=265 WHERE id_rubrica=14; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 14; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=265 WHERE id_rubrica=30; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 30; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Turistica - Nazionale' WHERE id_rubrica=266; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Turistica - Estera' WHERE id_rubrica=267; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Turismo & Cultura', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Istituzioni' WHERE id_rubrica=18; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Presidenti' WHERE id_rubrica=274; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Provincia di Venezia' WHERE id_rubrica=19; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Regione Veneto' WHERE id_rubrica=21; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Comuni' WHERE id_rubrica=20; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='ProLoco' WHERE id_rubrica=22; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 2; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='APT - Fornitori Servizi' WHERE id_rubrica=6; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Servizio medico - Pronto soccorso', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Servizio medico - Aziende ospedaliere', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='BUSSOLA - Richieste informazioni' WHERE id_rubrica=272; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ambiente', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Agricoltura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Commercio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Consorzi', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Artigiani', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Categorie consumatori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Industria', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ordini professionali', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Istruzione - Scuole', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Istruzione - Universit' + NCHAR(224) + ' & Accademie', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Autorit' + NCHAR(224) + ' giudiziarie', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Autorit' + NCHAR(224) + ' religiose', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Banche ed Istituti di Credito', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Parlamentari', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ristorazione', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Regione', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Provincia', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Comune', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Sindacati', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('LEO - Gruppo di lavoro', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('LEO - Sponsor', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('APT - Sostenitori', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Istituzioni fuori Veneto' WHERE id_rubrica=33; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 20' WHERE id_rubrica=277; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 21' WHERE id_rubrica=278; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 22' WHERE id_rubrica=279; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Territorio' WHERE id_rubrica=33; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Veneto' WHERE id_rubrica=33 "
CALL DB.Execute(sql, 111)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 112
'...........................................................................................
'modifiche dopo l'aggiornamento dell'oggetto in multilingue (vedi file .txt in 
'turismovenezia/oggetto ...)
'...........................................................................................
sql = " DROP VIEW dbo.ViewNotUtili; " + vbCrLf + _
	  "CREATE VIEW dbo.viewNotUtili" + vbCrLf + _
	  "AS" + vbCrLf + _
	  "SELECT     dbo.Not_util.id_UTIL, dbo.Not_util.Denom_util, dbo.Not_util.Indir1, dbo.Not_util.Cap, dbo.Not_util.Indir2, dbo.Not_util.Telef1, dbo.Not_util.Fax, " + vbCrLf + _
	  "                      dbo.Not_util.E_mail, dbo.Not_util.web, dbo.Not_util.zona, dbo.Not_util.Tipo, dbo.Not_util.[Local], dbo.Tipi_notutil.tipo_nome_eng, " + vbCrLf + _
	  "                      dbo.Tipi_notutil.tipo_nome_it, dbo.Tipi_notutil.tipo_nome_fra, dbo.Tipi_notutil.tipo_nome_ted, dbo.Tipi_notutil.tipo_nome_spa, " + vbCrLf + _
	  "                      dbo.Not_util.subzona" + vbCrLf + _
	  "FROM         dbo.Not_util INNER JOIN" + vbCrLf + _
	  "                      dbo.Tipi_notutil ON dbo.Not_util.Tipo = dbo.Tipi_notutil.id_tipoutil; " + vbCrLf

sql = sql & " ALTER TABLE PuntiStrat ADD den_fra nvarchar(50) NULL, den_spa nvarchar(50) NULL, den_ted nvarchar(50) NULL; "

sql = sql & " DROP VIEW dbo.viewDotazioni; " + _
			"CREATE VIEW dbo.viewDotazioni"+ vbCrLf + _
			"AS"+ vbCrLf + _
			"SELECT DISTINCT "+ vbCrLf + _
			"                      dbo.rel_ric_caratt.rel_id_struric, dbo.Caratt_TipiRic.caratt_nome_it, dbo.Caratt_TipiRic.caratt_sigla, dbo.Caratt_TipiRic.caratt_symb, "+ vbCrLf + _
			"                      dbo.Caratt_TipiRic.caratt_senum, dbo.Caratt_TipiRic.caratt_seprice, dbo.rel_ric_caratt.rel_valore, dbo.rel_ric_caratt.rel_testo_it, "+ vbCrLf + _
			"                      dbo.rel_ric_caratt.rel_testo_eng, dbo.rel_tipiStr_caratt.rel_ordine, dbo.rel_tipiStr_caratt.rel_id_tipo, dbo.Caratt_TipiRic.caratt_nome_eng, "+ vbCrLf + _
			"                      dbo.Caratt_TipiRic.caratt_nome_fra, dbo.Caratt_TipiRic.caratt_nome_ted, dbo.Caratt_TipiRic.caratt_nome_spa, dbo.rel_ric_caratt.rel_testo_fra, "+ vbCrLf + _
			"                      dbo.rel_ric_caratt.rel_testo_ted, dbo.rel_ric_caratt.rel_testo_spa"+ vbCrLf + _
			"FROM         dbo.Caratt_TipiRic INNER JOIN"+ vbCrLf + _
			"                      dbo.rel_ric_caratt ON dbo.Caratt_TipiRic.id_caratt = dbo.rel_ric_caratt.rel_id_caratt LEFT OUTER JOIN"+ vbCrLf + _
			"                      dbo.rel_tipiStr_caratt ON dbo.Caratt_TipiRic.id_caratt = dbo.rel_tipiStr_caratt.rel_id_caratt; "

sql = sql & " DROP VIEW dbo.viewServizi; "+ _
			"CREATE VIEW dbo.viewServizi "+ vbCrLf + _
			"AS"+ vbCrLf + _
			"SELECT     dbo.rel_Ric_Servizi.rel_ServRic_idRic, dbo.Servizi_TipiRic.serv_nome_it, dbo.Servizi_TipiRic.serv_symb, dbo.Servizi_TipiRic.serv_nome_eng, "+ vbCrLf + _
			"                      dbo.Servizi_TipiRic.serv_nome_ted, dbo.Servizi_TipiRic.serv_nome_fra, dbo.Servizi_TipiRic.serv_nome_spa"+ vbCrLf + _
			"FROM         dbo.rel_Serv_TipiRic INNER JOIN"+ vbCrLf + _
			"                      dbo.rel_Ric_Servizi ON dbo.rel_Serv_TipiRic.rel_tipiRicServ_id = dbo.rel_Ric_Servizi.rel_relServ_Ric INNER JOIN"+ vbCrLf + _
			"                      dbo.Servizi_TipiRic ON dbo.rel_Serv_TipiRic.rel_tRicServ_idServ = dbo.Servizi_TipiRic.id_serv_tipiric; "
			
sql = sql & " UPDATE tb_tipo_appunti SET descrizione = 'Luoghi' WHERE id_tipo_appunti=1; " + _
			" UPDATE tb_tipo_appunti SET descrizione = 'Eventi' WHERE id_tipo_appunti=2; " + _
			" UPDATE tb_tipo_appunti SET descrizione = 'Strutture ricettive' WHERE id_tipo_appunti=3; " + _
			" UPDATE tb_tipo_appunti SET descrizione = 'Locali e servizi' WHERE id_tipo_appunti=4; " + _
			" UPDATE tb_tipo_appunti SET descrizione = 'Notizie utili' WHERE id_tipo_appunti=5; " + _
			" UPDATE tb_tipo_appunti SET descrizione = 'Spiagge' WHERE id_tipo_appunti=6; " + _
			" INSERT INTO tb_tipo_appunti (descrizione) VALUES ('Pagine'); "

CALL DB.Execute(sql, 112)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 113
'...........................................................................................
'copia valori campi in inglese su altre lingue dove non NULL o vuoto nelle tabelle di ricerca
'...........................................................................................
dim campoEng, campo, tab
tab = "zone_struRic"
campoEng = "zona_nome_eng"
campo = "zona_nome_fra"
sql = " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "zona_nome_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "zona_nome_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "tipi_ls"
campoEng = "tipo_nome_eng"
campo = "tipo_nome_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "tipo_nome_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "tipo_nome_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "tipi_notutil"
campoEng = "tipo_nome_eng"
campo = "tipo_nome_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "tipo_nome_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "tipo_nome_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "sottotipi_notutil"
campoEng = "sottip_nome_eng"
campo = "sottip_nome_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "sottip_nome_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "sottip_nome_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "sottotipi_locserv"
campoEng = "sottip_nome_eng"
campo = "sottip_nome_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "sottip_nome_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "sottip_nome_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "eventispeciali"
campoEng = "ev_spec_eng"
campo = "ev_spec_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "ev_spec_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "ev_spec_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "categorieeventi"
campoEng = "desc_eng"
campo = "desc_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "desc_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "desc_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

tab = "tipi_ric"
campoEng = "denominazione_eng"
campo = "denominazione_fra"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "denominazione_spa"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "
campo = "denominazione_ted"
sql = sql + " UPDATE "& tab &" SET "& campo &" = "& campoEng &" WHERE IsNull("& campo &", '') = ''; "

CALL DB.Execute(sql, 113)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 114
'...........................................................................................
'cancella e ricrea le principali viste x aggiornarle da eventuali cambiamenti di tabelle su
'cui viene eseguito un SELECT *
'...........................................................................................
sql = " DROP VIEW dbo.viewAllNotUtili; "+ _
	  "CREATE VIEW dbo.viewAllNotUtili"+ vbCrLf + _
	  "AS"+ vbCrLf + _
	  "SELECT     *"+ vbCrLf + _
	  "FROM         dbo.Not_util INNER JOIN"+ vbCrLf + _
	  "                      dbo.Tipi_notutil ON dbo.Not_util.Tipo = dbo.Tipi_notutil.id_tipoutil; "+ vbCrLf + _
	  " DROP VIEW dbo.viewLocaliEServizi; "+ vbCrLf + _
	  "CREATE VIEW dbo.viewLocalieServizi "+ vbCrLf + _
	  "AS"+ vbCrLf + _
	  "SELECT dbo.LocalieServizi.*, dbo.Tipi_LS.tipo_nome_it, dbo.Tipi_LS.tipo_nome_eng"+ vbCrLf + _
	  "FROM dbo.LocalieServizi INNER JOIN dbo.Tipi_LS ON dbo.LocalieServizi.Tipo = dbo.Tipi_LS.id_tipoutil; "
	  
CALL DB.Execute(sql, 114)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 115
'...........................................................................................
'svuota indirizzario per errore di import dati
'...........................................................................................
sql = "DELETE FROM tb_indirizzario; " &_
	  "DELETE FROM tb_rubriche; "
CALL DB.Execute(sql, 115)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 116
'...........................................................................................
'Aggiunge rubriche di sistema per sincronizzazione con dati del portale
'...........................................................................................
sql = " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Strutture ricettive', 'Stru_ric', 'Stru_ric', 1, 1) " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Spiagge', 'Spiagge', NULL, 1, 1) " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Luoghi', 'Luoghi', 'Luoghi', 1, 1) " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Locali & Servizi', 'LocaliEServizi', 'LocaliEServizi', 1, 1) " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, locked_rubrica, rubrica_esterna) " & _
	  " VALUES ('Notizie utili', 'Not_Util', 'Not_Util', 1, 1) " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Strutture ricettive - ' + Denominazione_it, 	'Stru_ric', 		'Tipi_ric', IDTipo, 1, 1 FROM Tipi_Ric; " & _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Locali & Servizi - ' + tipo_nome_it, 			'LocaliEServizi', 	'Tipi_LS', id_tipoUtil, 1, 1 FROM Tipi_LS; " &_
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Notizie utili - ' + tipo_nome_it, 				'Not_Util', 		'Tipi_notUtil', id_tipoUtil, 1, 1 FROM Tipi_NotUtil; " & _
  	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  " 	SELECT 'Luoghi - ' + tipo_luogo, 						'Luoghi', 			'TipoLuoghi', IDL, 1, 1 FROM TipoLuoghi WHERE Visibile=1; " & _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " &_
	  "		SELECT 'Notizie utili - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'Not_Util', 'SottoTipi_NotUtil', id_sottipo, 1,1 " &_
	  "			FROM SottoTipi_NotUtil INNER JOIN Tipi_notUtil ON SottoTipi_NotUtil.ref_tipo=tipi_NotUtil.id_tipoUtil ;" &_
	  "	INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & _
	  "		SELECT 'Locali & Servizi - ' + tipo_nome_it + ' - ' + sottip_nome_it, 'LocaliEServizi', 'SottoTipi_LocServ', id_sottipo, 1,1 " &_
	  "			FROM SottoTipi_LocServ INNER JOIN Tipi_LS ON SottoTipi_LocServ.ref_tipo=Tipi_LS.id_tipoUtil ;" &_
	  "	INSERT INTO tb_rel_gruppirubriche (id_dellarubrica, id_gruppo_assegnato) " &_
	  " 	SELECT tb_rubriche.id_rubrica, tb_gruppi.id_gruppo FROM tb_rubriche, tb_gruppi "
CALL DB.Execute(sql, 116)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 117
'...........................................................................................
'importa dati da database indirizzario (richiede dbindirizzario.mdb in cartella database
'...........................................................................................
'AGGIORNAMENTO rimosso dal framework per rimozione applicativo AptAdmin il 03/12/2007
'<!--#include file="subscripts/Update_govenice_117_Inport_Indirizzario.asp"-->
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 117)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 118
'...........................................................................................
'completa sincronizzazione record vari con contati (LUOGHI, strutture Ricettive, Notizie utili)
'...........................................................................................
'AGGIORNAMENTO rimosso dal framework per rimozione applicativo AptAdmin il 03/12/2007
'<!--#include file="subscripts/Update_govenice_118_Completa_syncro.asp"-->
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 118)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 119
'...........................................................................................
'pulisce valori numeri non validi
'...........................................................................................
sql = " DELETE FROM tb_ValoriNumeri WHERE ISNULL(ValoreNumero, '')=''; "	  
CALL DB.Execute(sql, 119)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 120
'...........................................................................................
'esegue modifiche alle rubriche
'...........................................................................................
sql = " UPDATE tb_rubriche SET nome_rubrica='APT - Provincia di Venezia' WHERE id_rubrica=23; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='APT - Fuuori Provincia di Venezia' WHERE id_rubrica=269; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Aziende & Societ' + NCHAR(224) + ' varie' WHERE id_rubrica=9; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Autorit' + NCHAR(224) + ' ambito APT' WHERE id_rubrica=263; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Agenzie di viaggio & Tour Operator', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Cultura Italia - Biblioteche', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Musei' WHERE id_rubrica=24; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Musica' WHERE id_rubrica=26; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Italia - Enti & Societ' + NCHAR(224) WHERE id_rubrica=34; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Cultura Estero' WHERE id_rubrica=7; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=34 WHERE id_rubrica=27; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 27; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Venezia / Mestre - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Venezia / Mestre - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Riviera del Brenta - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Riviera del Brenta - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Terra dei Tiepolo - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Terra dei Tiepolo - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Cavallino - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Cavallino - Assessori Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Quarto d''Altino - Sindaci', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comuni - Quarto d''Altino - Assessori Turismo & Cultura', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Assessori turismo ambito APT' WHERE id_rubrica=5; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da controllare - Sindaci' WHERE id_rubrica=4; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Assessori Turismo fuori APT' WHERE id_rubrica=268; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consigli di Quartiere & Municipalit' + NCHAR(224) WHERE id_rubrica=17; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consolati & Ambasciate Italiane all''estero' WHERE id_rubrica=12; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Consolati & Ambasciate Estere in Italia' WHERE id_rubrica=11; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Locale' WHERE id_rubrica=3; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=3 WHERE id_rubrica=13; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 13; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Nazionale' WHERE id_rubrica=264; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=264 WHERE id_rubrica=28; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=264 WHERE id_rubrica=29; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 28; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 29; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Generale - Estera' WHERE id_rubrica=265; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=265 WHERE id_rubrica=14; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 14; " &_
	  " UPDATE rel_rub_ind SET id_rubrica=265 WHERE id_rubrica=30; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 30; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Turistica - Nazionale' WHERE id_rubrica=266; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='MEDIA - Turistica - Estera' WHERE id_rubrica=267; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Regione Veneto - Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Provincia di Venezia - Turismo & Cultura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Giunta & Consiglio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Amministratori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Comune di Venezia - Turismo & Cultura', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Istituzioni' WHERE id_rubrica=18; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Presidenti' WHERE id_rubrica=274; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Provincia di Venezia' WHERE id_rubrica=19; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Regione Veneto' WHERE id_rubrica=21; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Comuni' WHERE id_rubrica=20; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='ProLoco' WHERE id_rubrica=22; " &_
	  " DELETE FROM tb_rubriche WHERE id_rubrica = 2; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='APT - Fornitori Servizi' WHERE id_rubrica=6; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Servizio medico - Pronto soccorso', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Servizio medico - Aziende ospedaliere', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='BUSSOLA - Richieste informazioni' WHERE id_rubrica=272; " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ambiente', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Agricoltura', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Commercio', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Consorzi', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Artigiani', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Categorie consumatori', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Industria', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ordini professionali', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Istruzione - Scuole', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Istruzione - Universit' + NCHAR(224) + ' & Accademie', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Autorit' + NCHAR(224) + ' giudiziarie', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Autorit' + NCHAR(224) + ' religiose', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Banche ed Istituti di Credito', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Parlamentari', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Ristorazione', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Regione', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Provincia', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Segreterie partiti - Comune', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('Sindacati', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('LEO - Gruppo di lavoro', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('LEO - Sponsor', 0, 0); " &_
	  " INSERT INTO tb_rubriche(nome_Rubrica, locked_rubrica, rubrica_esterna) VALUES ('APT - Sostenitori', 0, 0); " &_
	  " UPDATE tb_rubriche SET nome_rubrica='Istituzioni fuori Veneto' WHERE id_rubrica=33; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 20' WHERE id_rubrica=277; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 21' WHERE id_rubrica=278; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='LEO - Si parla di voi - 22' WHERE id_rubrica=279; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Territorio' WHERE id_rubrica=33; " &_
	  " UPDATE tb_rubriche SET nome_rubrica='_Da Controllare - Veneto' WHERE id_rubrica=33 "
CALL DB.Execute(sql, 120)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 121
'...........................................................................................
'cancella e ricrea vista eventi per luogo
'...........................................................................................
sql = " ALTER VIEW dbo.viewEventiPerLuogo AS " + _
	  " SELECT Eventi.*, dbo.Luoghi.Zona, dbo.Luoghi.subZona " + _
	  "		FROM dbo.Eventi INNER JOIN dbo.doveAccade ON dbo.Eventi.ID = dbo.doveAccade.id_evento " + _
	  "		INNER JOIN dbo.Luoghi ON dbo.doveAccade.id_luogo = dbo.Luoghi.ID "	  
CALL DB.Execute(sql, 121)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 122
'...........................................................................................
'allarga dimensione campo categoria
'...........................................................................................
sql = " ALTER TABLE Stru_ric ALTER COLUMN categoria nvarchar(20) "	  
CALL DB.Execute(sql, 122)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 123
'...........................................................................................
'crea tabella per tracciatura esecuzione aggiornamenti globali del webservice
'...........................................................................................
sql = " CREATE TABLE dbo.Stru_Ric_WS_Config ( " + _
	  " 	ultimo_Aggiornamento smalldatetime NULL " + _
	  "		); " + _
	  " INSERT INTO Stru_Ric_WS_Config (ultimo_aggiornamento) VALUES(GETDATE() - 30)"
CALL DB.Execute(sql, 123)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 124
'...........................................................................................
'modifica vista per correzione problema venicesystem (mappe)
'...........................................................................................
sql = " DROP VIEW vwp_disponibilita ; " + _
	  " CREATE VIEW dbo.vwp_disponibilita AS " + vbCrlf + _
	  "		SELECT TOP 100 PERCENT disp_id, disp_data, disp_apertura, disp_numero_max, " + vbCrlf + _
	  "			disp_camere_libere, disp_id_albergo, disp_id_tipo_camera, disp_prezzo , tbp_tipiCamera.*, " + vbCrlf + _
	  "			(dbo.Stru_ric.ID_albergo) As str_id, collocazione, linkmappe, VIEW_alberghi.* " + vbCrlf + _
	  "		FROM tbp_disponibilita INNER JOIN VIEW_alberghi " + vbCrlf + _
	  "			ON tbp_disponibilita.disp_id_albergo = VIEW_alberghi.id_albergo " + vbCrlf + _
	  "			AND tbp_disponibilita.disp_data >=( GETDATE() -1) + VIEW_alberghi.preavviso_albergo " + vbCrlf + _
	  "			INNER JOIN tbp_tipiCamera ON tbp_disponibilita.disp_id_tipo_camera = tbp_tipiCamera.id_tipoCamera " + vbCrlf + _
	  "			LEFT OUTER JOIN Stru_ric ON VIEW_alberghi.codice_albergo = Stru_ric.RegCode " + vbCrlf + _
	  "		WHERE (tbp_disponibilita.disp_apertura = 1) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_camere_libere>0) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_data >= (GETDATE() - 1) + VIEW_alberghi.preavviso_albergo) " + vbCrlf + _
	  "			ORDER BY VIEW_alberghi.PREZZO_MEDIO_PL, tbp_disponibilita.disp_id_albergo, tbp_disponibilita.disp_id_tipo_camera, tbp_disponibilita.disp_data "
CALL DB.Execute(sql, 124)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 125
'...........................................................................................
'modifica vista per correzione problema venicesystem (mappe)
'...........................................................................................
sql = " DROP VIEW vwp_disponibilita ; " + _
	  " CREATE VIEW dbo.vwp_disponibilita AS " + vbCrlf + _
	  "		SELECT TOP 100 PERCENT disp_id, disp_data, disp_apertura, disp_numero_max, " + vbCrlf + _
	  "			disp_camere_libere, disp_id_albergo, disp_id_tipo_camera, disp_prezzo , tbp_tipiCamera.*, " + vbCrlf + _
	  "			(dbo.Stru_ric.ID_albergo) As str_id, (linkmappe) AS collocazione, VIEW_alberghi.* " + vbCrlf + _
	  "		FROM tbp_disponibilita INNER JOIN VIEW_alberghi " + vbCrlf + _
	  "			ON tbp_disponibilita.disp_id_albergo = VIEW_alberghi.id_albergo " + vbCrlf + _
	  "			AND tbp_disponibilita.disp_data >=( GETDATE() -1) + VIEW_alberghi.preavviso_albergo " + vbCrlf + _
	  "			INNER JOIN tbp_tipiCamera ON tbp_disponibilita.disp_id_tipo_camera = tbp_tipiCamera.id_tipoCamera " + vbCrlf + _
	  "			LEFT OUTER JOIN Stru_ric ON VIEW_alberghi.codice_albergo = Stru_ric.RegCode " + vbCrlf + _
	  "		WHERE (tbp_disponibilita.disp_apertura = 1) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_camere_libere>0) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_data >= (GETDATE() - 1) + VIEW_alberghi.preavviso_albergo) " + vbCrlf + _
	  "			ORDER BY VIEW_alberghi.PREZZO_MEDIO_PL, tbp_disponibilita.disp_id_albergo, tbp_disponibilita.disp_id_tipo_camera, tbp_disponibilita.disp_data "
CALL DB.Execute(sql, 125)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 126
'...........................................................................................
'modifica vista per correzione problema venicesystem (mappe)
'...........................................................................................
sql = " DROP VIEW vwp_disponibilita ; " + _
	  " CREATE VIEW dbo.vwp_disponibilita AS " + vbCrlf + _
	  "		SELECT TOP 100 PERCENT disp_id, disp_data, disp_apertura, disp_numero_max, " + vbCrlf + _
	  "			disp_camere_libere, disp_id_albergo, disp_id_tipo_camera, disp_prezzo , tbp_tipiCamera.*, " + vbCrlf + _
	  "			(dbo.Stru_ric.ID_albergo) As str_id, collocazione, linkmappe, VIEW_alberghi.* " + vbCrlf + _
	  "		FROM tbp_disponibilita INNER JOIN VIEW_alberghi " + vbCrlf + _
	  "			ON tbp_disponibilita.disp_id_albergo = VIEW_alberghi.id_albergo " + vbCrlf + _
	  "			AND tbp_disponibilita.disp_data >=( GETDATE() -1) + VIEW_alberghi.preavviso_albergo " + vbCrlf + _
	  "			INNER JOIN tbp_tipiCamera ON tbp_disponibilita.disp_id_tipo_camera = tbp_tipiCamera.id_tipoCamera " + vbCrlf + _
	  "			LEFT OUTER JOIN Stru_ric ON VIEW_alberghi.codice_albergo = Stru_ric.RegCode " + vbCrlf + _
	  "		WHERE (tbp_disponibilita.disp_apertura = 1) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_camere_libere>0) " + vbCrlf + _
	  "			AND (tbp_disponibilita.disp_data >= (GETDATE() - 1) + VIEW_alberghi.preavviso_albergo) " + vbCrlf + _
	  "			ORDER BY VIEW_alberghi.PREZZO_MEDIO_PL, tbp_disponibilita.disp_id_albergo, tbp_disponibilita.disp_id_tipo_camera, tbp_disponibilita.disp_data "
CALL DB.Execute(sql, 126)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 127
'...........................................................................................
'allunga campi su tabella spiagge
'...........................................................................................
sql = " ALTER TABLE Spiagge ALTER COLUMN Denominazione nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Indir1_estate nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Indir2_inverno nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Localita nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Comune nvarchar (50) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Tel1_estate nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Fax_estate nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Tel2_inverno nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Fax_inverno nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Email nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Web nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Apertura_stagionale nvarchar (250) NULL ;" + _
	  " ALTER TABLE Spiagge ALTER COLUMN Orari nvarchar (250) NULL "
CALL DB.Execute(sql, 127)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 128
'...........................................................................................
'inserisco tabelle per NEXTCRM
'...........................................................................................
sql = "CREATE TABLE [dbo].[tb_descrittori] ( " + vbCrlf + _
	  "[descr_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[descr_nome] [nvarchar] (50) NULL , " + vbCrlf + _
	  "[descr_tipo] [smallint] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[tb_tipologie] ( " + vbCrlf + _
	  "[tipo_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[tipo_nome] [nvarchar] (50) NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[rel_tipologie_descrittori] ( " + vbCrlf + _
	  "[rtd_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[rtd_tipologia_id] [int] NULL , " + vbCrlf + _
	  "[rtd_descrittore_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[tb_pratiche] ( " + vbCrlf + _
	  "[pra_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[pra_codice] [nvarchar] (50) NULL , " + vbCrlf + _
	  "[pra_nome] [nvarchar] (255) NULL , " + vbCrlf + _
	  "[pra_dataI] [smalldatetime] NULL , " + vbCrlf + _
	  "[pra_dataUM] [smalldatetime] NULL , " + vbCrlf + _
	  "[pra_dataA] [smalldatetime] NULL , " + vbCrlf + _
	  "[pra_archiviata] [bit] NULL , " + vbCrlf + _
	  "[pra_note] [ntext] NULL , " + vbCrlf + _
	  "[pra_pubblica] [bit] NULL , " + vbCrlf + _
	  "[pra_cliente_id] [int] NULL , " + vbCrlf + _
	  "[pra_creatore_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_default_gruppi] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_gruppo_id] [int] NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_default_utenti] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_utente_id] [int] NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_pratiche_gruppi] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_gruppo_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_pratiche_utenti] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_utente_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[tb_attivita] ( " + vbCrlf + _
	  "[att_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[att_oggetto] [nvarchar] (255) NULL , " + vbCrlf + _
	  "[att_testo] [ntext] NULL , " + vbCrlf + _
	  "[att_note] [ntext] NULL , " + vbCrlf + _
	  "[att_dataCrea] [smalldatetime] NULL , " + vbCrlf + _
	  "[att_dataChiusa] [smalldatetime] NULL , " + vbCrlf + _
	  "[att_dataS] [smalldatetime] NULL , " + vbCrlf + _
	  "[att_priorita] [bit] NULL , " + vbCrlf + _
	  "[att_conclusa] [bit] NULL , " + vbCrlf + _
	  "[att_pubblica] [bit] NULL , " + vbCrlf + _
	  "[att_eredita] [bit] NULL , " + vbCrlf + _
	  "[att_sistema] [bit] NULL , " + vbCrlf + _
	  "[att_domanda_id] [int] NULL , " + vbCrlf + _
	  "[att_mittente_id] [int] NULL , " + vbCrlf + _
	  "[att_pratica_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[tb_documenti] ( " + vbCrlf + _
	  "[doc_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[doc_nome] [nvarchar] (255) NULL , " + vbCrlf + _
	  "[doc_path] [nvarchar] (255) NULL , " + vbCrlf + _
	  "[doc_dataC] [smalldatetime] NULL , " + vbCrlf + _
	  "[doc_pubblica] [bit] NULL , " + vbCrlf + _
      "[doc_eredita] [bit] NULL , " + vbCrlf + _
	  "[doc_note] [ntext] NULL , " + vbCrlf + _
	  "[doc_tipologia_id] [int] NULL , " + vbCrlf + _
	  "[doc_pratica_id] [int] NULL , " + vbCrlf + _
	  "[doc_creatore_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_attivita_gruppi] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_gruppo_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_attivita_utenti] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_utente_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_documenti_gruppi] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_gruppo_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[al_documenti_utenti] ( " + vbCrlf + _
	  "[al_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[al_tipo_id] [int] NULL , " + vbCrlf + _
	  "[al_utente_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[rel_documenti_descrittori] ( " + vbCrlf + _
	  "[rdd_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[rdd_valore] [nvarchar] (255) NULL , " + vbCrlf + _
	  "[rdd_documento_id] [int] NULL , " + vbCrlf + _
	  "[rdd_descrittore_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "CREATE TABLE [dbo].[tb_allegati] ( " + vbCrlf + _
	  "[all_id] [int] IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
	  "[all_attivita_id] [int] NULL , " + vbCrlf + _
	  "[all_documento_id] [int] NULL  " + vbCrlf + _
	  ") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_descrittori] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_descrittori] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[descr_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_tipologie] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_tipologie] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[tipo_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[rel_tipologie_descrittori] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_tipologie_descrittori] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rtd_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_pratiche] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_pratiche] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_default_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_al_default_gruppi] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		")  ON [PRIMARY]  " + vbCrlf + _
	  "; " + vbCrlf + _
 	  "ALTER TABLE [dbo].[al_default_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_al_default] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_pratiche_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_pratica_gruppo] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_pratiche_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_pratiche_utenti] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_attivita] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_messaggi] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[att_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_documenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_documenti] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[doc_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_attivita_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_al_messaggi_gruppi] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_attivita_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_al_attivita_utenti] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_documenti_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_documenti_gruppi] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_documenti_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_documenti_utenti] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[rel_documenti_descrittori] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_rel_documenti_descrittori] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rdd_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_allegati] ADD  " + vbCrlf + _
	  "CONSTRAINT [PK_tb_allegati] PRIMARY KEY  CLUSTERED  " + vbCrlf + _
		"( " + vbCrlf + _
			"[all_id] " + vbCrlf + _
		") " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_pratiche] ADD  " + vbCrlf + _
	  "CONSTRAINT [DF_tb_pratiche_pra_archiviata] DEFAULT (0) FOR [pra_archiviata] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_attivita] ADD  " + vbCrlf + _
	  "CONSTRAINT [DF_tb_attivita_att_priorita] DEFAULT (0) FOR [att_priorita], " + vbCrlf + _
	  "CONSTRAINT [DF_tb_messaggi_msg_conclusa] DEFAULT (0) FOR [att_conclusa], " + vbCrlf + _
	  "CONSTRAINT [DF_tb_attivita_att_sistema] DEFAULT (0) FOR [att_sistema] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_documenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [DF_tb_documenti_doc_pubblica] DEFAULT (0) FOR [doc_pubblica], " + vbCrlf + _
	  "CONSTRAINT [DF_tb_documenti_doc_eredita] DEFAULT (1) FOR [doc_eredita] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[rel_tipologie_descrittori] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_rel_tipologie_descrittori_tb_descrittori] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rtd_descrittore_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_descrittori] ( " + vbCrlf + _
			"[descr_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_rel_tipologie_descrittori_tb_tipologie] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rtd_tipologia_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_tipologie] ( " + vbCrlf + _
			"[tipo_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_pratiche] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_tb_pratiche_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[pra_creatore_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		"), " + vbCrlf + _
	  "CONSTRAINT [FK_tb_pratiche_tb_Indirizzario] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[pra_cliente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_Indirizzario] ( " + vbCrlf + _
			"[IDElencoIndirizzi] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "alter table [dbo].[tb_pratiche] nocheck constraint [FK_tb_pratiche_tb_admin] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_default_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_default_gruppi_tb_gruppi] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_gruppo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_gruppi] ( " + vbCrlf + _
			"[id_Gruppo] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_default_gruppi_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_default_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_default_utenti_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_utente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_default_utenti_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_pratiche_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_pratiche_gruppi_tb_gruppi] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_gruppo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_gruppi] ( " + vbCrlf + _
			"[id_Gruppo] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	   "CONSTRAINT [FK_al_pratiche_gruppi_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_pratiche_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_pratiche_utenti_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_utente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_pratiche_utenti_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_attivita] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_tb_attivita_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[att_mittente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		"), " + vbCrlf + _
	  "CONSTRAINT [FK_tb_messaggi_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[att_pratica_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "alter table [dbo].[tb_attivita] nocheck constraint [FK_tb_attivita_tb_admin] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_documenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_tb_documenti_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[doc_creatore_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		"), " + vbCrlf + _
	  "CONSTRAINT [FK_tb_documenti_tb_pratiche] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[doc_pratica_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_pratiche] ( " + vbCrlf + _
			"[pra_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_tb_documenti_tb_tipologie] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[doc_tipologia_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_tipologie] ( " + vbCrlf + _
			"[tipo_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "alter table [dbo].[tb_documenti] nocheck constraint [FK_tb_documenti_tb_admin] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_attivita_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_messaggi_gruppi_tb_attivita] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_attivita] ( " + vbCrlf + _
			"[att_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_messaggi_gruppi_tb_gruppi] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_gruppo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_gruppi] ( " + vbCrlf + _
			"[id_Gruppo] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_attivita_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_attivita_utenti_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_utente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_attivita_utenti_tb_attivita] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_attivita] ( " + vbCrlf + _
			"[att_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_documenti_gruppi] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_documenti_gruppi_tb_documenti] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_documenti] ( " + vbCrlf + _
			"[doc_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_documenti_gruppi_tb_gruppi] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_gruppo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_gruppi] ( " + vbCrlf + _
			"[id_Gruppo] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[al_documenti_utenti] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_al_documenti_utenti_tb_admin] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_utente_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_admin] ( " + vbCrlf + _
			"[id_admin] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_al_documenti_utenti_tb_documenti] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[al_tipo_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_documenti] ( " + vbCrlf + _
			"[doc_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[rel_documenti_descrittori] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_rel_documenti_descrittori_tb_descrittori] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rdd_descrittore_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_descrittori] ( " + vbCrlf + _
			"[descr_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrlf + _
	  "CONSTRAINT [FK_rel_documenti_descrittori_tb_documenti] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[rdd_documento_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_documenti] ( " + vbCrlf + _
			"[doc_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE [dbo].[tb_allegati] ADD  " + vbCrlf + _
	  "CONSTRAINT [FK_tb_allegati_tb_documenti] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[all_documento_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_documenti] ( " + vbCrlf + _
			"[doc_id] " + vbCrlf + _
		"), " + vbCrlf + _
	  "CONSTRAINT [FK_tb_allegati_tb_messaggi] FOREIGN KEY  " + vbCrlf + _
		"( " + vbCrlf + _
			"[all_attivita_id] " + vbCrlf + _
		") REFERENCES [dbo].[tb_attivita] ( " + vbCrlf + _
			"[att_id] " + vbCrlf + _
		") ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrlf + _
	  "; " + vbCrlf + _
	  "alter table [dbo].[tb_allegati] nocheck constraint [FK_tb_allegati_tb_documenti] " + vbCrlf + _
	  "; " + vbCrlf + _
	  "ALTER TABLE tb_indirizzario ADD " + vbCrlf + _
	  "[PraticaPrefisso] [char] (5) NULL , " + vbCrlf + _
	  "[PraticaCount] [int] NULL DEFAULT (0)"

CALL DB.Execute(sql, 128)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 129
'...........................................................................................
'aggiunge il permesso COM_POWER per il CRM
'...........................................................................................
sql = "UPDATE tb_siti SET sito_p3='COM_POWER' WHERE id_sito=3"
CALL DB.Execute(sql, 129)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 130
'...........................................................................................
'aggiunge il permesso COM_POWER per il CRM
'...........................................................................................
sql = "ALTER TABLE Eventi ALTER COLUMN ev_tel_org char (250) NULL ;"
CALL DB.Execute(sql, 130)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 131
'...........................................................................................
'aggiunge campi luogo di nascita e codice fiscale all'indirizzario
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD " + _
	  " LuogoNascita nvarchar(255) NULL ," + _
	  " CF nvarchar(16) NULL "
CALL DB.Execute(sql, 131)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 132
'...........................................................................................
'aggiornamento per compensare mancanza campi su tabella email
'...........................................................................................
sql = "ALTER TABLE tb_email " &_
		"ADD email_mime nvarchar(50) NULL, " &_
		"	email_From nvarchar(250) NULL" 
CALL DB.Execute(sql, 132)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 133
'...........................................................................................
'cambia proprietario oggetti
'...........................................................................................
sql = "EXEC dbo.sp_changeobjectowner 'Stru_Ric_WS_Config', 'dbo';"  + _
 	  "EXEC dbo.sp_changeobjectowner 'Delete_tbp_dettagli', 'dbo';"  + _
	  "EXEC dbo.sp_changeobjectowner 'Delete_tbp_pacchetti', 'dbo';"  + _
	  "EXEC dbo.sp_changeobjectowner 'Delete_tbp_alberghi', 'dbo'"
CALL DB.Execute(sql, 133)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 134
'...........................................................................................
'corregge sincronizzazione record vari con contati (LUOGHI, strutture Ricettive, Notizie utili)
'...........................................................................................
'AGGIORNAMENTO rimosso dal framework per rimozione applicativo AptAdmin il 03/12/2007
'<!--#include file="subscripts/Update_govenice_134_Completa_syncro.asp"-->
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 135)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 135
'...........................................................................................
'aggiunge campo per la relazione tra contatti dell'indirizario
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD " + vbCrlf + _
	  "[cntRel] [int] NULL"
CALL DB.Execute(sql, 135)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 136
'...........................................................................................
'allarga campo del codice regionale su tabella strutture ricettive
'...........................................................................................
sql = "ALTER TABLE Stru_ric ALTER COLUMN RegCode nvarchar(12)"
CALL DB.Execute(sql, 136)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 137
'...........................................................................................
'allarga campo del codice regionale su tabella strutture ricettive
'...........................................................................................
sql = "ALTER TABLE Stru_ric ALTER COLUMN foto_int nvarchar(200); " & vbcrlf & _
	  "ALTER TABLE Stru_ric ALTER COLUMN foto_est nvarchar(200); " & vbcrlf & _
	  "ALTER TABLE Stru_ric ADD foto_int_big nvarchar(200), foto_est_big nvarchar(200);"
CALL DB.Execute(sql, 137)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 138
'...........................................................................................
'aggiunge il campo default per le pratiche causa gestione documenti
'...........................................................................................
sql = "ALTER TABLE tb_pratiche ADD " + vbCrlf + _
	  "[pra_default] [bit] NULL"
CALL DB.Execute(sql, 138)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 139
'...........................................................................................
'allarga campo del codice regionale su tabella strutture ricettive
'...........................................................................................
sql = "ALTER TABLE Stru_ric ALTER COLUMN Denominazione nvarchar(150)"
CALL DB.Execute(sql, 139)
'******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 140
'...........................................................................................
'allarga campo del codice regionale su tabella strutture ricettive
'...........................................................................................
sql = "ALTER VIEW dbo.viewLocalieServizi AS " & _
	  "SELECT     dbo.LocalieServizi.*, " & _
	  " dbo.Tipi_LS.tipo_nome_it, dbo.Tipi_LS.tipo_nome_eng, dbo.Tipi_LS.tipo_nome_fra, dbo.Tipi_LS.tipo_nome_ted, dbo.Tipi_LS.tipo_nome_spa "  & _
	  " FROM         dbo.LocalieServizi INNER JOIN dbo.Tipi_LS ON dbo.LocalieServizi.Tipo = dbo.Tipi_LS.id_tipoutil "
	  CALL DB.Execute(sql, 140)
'******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 141
'...........................................................................................
'Toglie campo flag per pratica default e toglie relazioni tra attivita' e pratiche 
' e documenti e pratiche
'...........................................................................................
sql = "ALTER TABLE tb_pratiche DROP COLUMN pra_default; " & vbCrLf & _
	  "ALTER TABLE tb_attivita DROP CONSTRAINT FK_tb_messaggi_tb_pratiche; " & vbCrLf & _
	  "ALTER TABLE tb_documenti DROP CONSTRAINT FK_tb_documenti_tb_pratiche " & vbCrLf
CALL DB.Execute(sql, 141)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 142
'...........................................................................................
'aggiunge gestione lingua su nextCom
'...........................................................................................
sql = "ALTER TABLE tb_indirizzario ADD lingua varchar(2); " & vbcrLF & _
	  "UPDATE tb_indirizzario SET lingua = 'it' "
CALL DB.Execute(sql, 142)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 143
'...........................................................................................
'aggiunge gestione lingua su nextCom
'...........................................................................................
sql = "CREATE TABLE dbo.tb_cnt_lingue (" & vbCrLF & _
	  "		lingua_codice varchar(2) NOT NULL CONSTRAINT PK_tb_cnt_lingue PRIMARY KEY, " & vbCrLF & _
	  "		lingua_nome_IT varchar(20) NULL, " & vbCrLf & _
	  "		lingua_nome varchar(20) NULL " & vbCRLF & _
	  "		); " & vbCRLF & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('it', 'Italiano', 'Italiano'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('en', 'Inglese', 'English'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('fr', 'Francese', 'Français'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('de', 'Tedesco', 'Deutsch'); " & vbCrLf & _
	  "INSERT INTO tb_cnt_lingue (lingua_codice, lingua_nome_it, lingua_nome) VALUES ('es', 'Spagnolo', 'Español'); " & vbCrLf & _
	  "ALTER TABLE tb_indirizzario ADD CONSTRAINT FK_tb_indirizzario__tb_cnt_lingue " & vbCrLf &_
	  "		FOREIGN KEY (lingua) REFERENCES tb_cnt_lingue(lingua_codice) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 143)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 144
'...........................................................................................
'aggiunge gestione parametri di installazione / funzionamento alle applicazioni del next-passport
'...........................................................................................
sql = " CREATE TABLE dbo.tb_siti_parametri (" & _
	  "		par_id INT IDENTITY (1, 1) NOT NULL, " & _
	  "		par_key varchar(50) NOT NULL, " & _
	  "		par_value varchar(250) NULL, " & _
	  "		par_sito_id INT NOT NULL " & _
	  "		); " & _
	  "	ALTER TABLE tb_siti_parametri ADD CONSTRAINT PK_tb_siti_parametri PRIMARY KEY NONCLUSTERED (par_id); " &_
	  " ALTER TABLE tb_siti_parametri ADD CONSTRAINT FK_tb_siti_parametri__tb_siti " & _
	  "		FOREIGN KEY (par_sito_id) REFERENCES tb_siti(id_sito) " & _
	  "		ON UPDATE CASCADE ON DELETE CASCADE "
CALL DB.Execute(sql, 144)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 145
'...........................................................................................
'aggiunge campi per gestione informazioni disabili nella tabella Luoghi
'...........................................................................................
sql = "ALTER TABLE Luoghi ADD "& _
	  "		info_disabili_it NTEXT NULL, "& _
	  "		info_disabili_en NTEXT NULL, "& _
	  "		info_disabili_es NTEXT NULL, "& _
	  "		info_disabili_de NTEXT NULL, "& _
	  "		info_disabili_fr NTEXT NULL, "& _
	  "		info_disabili_vis BIT NULL;" & _
	  "UPDATE Luoghi SET info_disabili_vis=0"
CALL DB.Execute(sql, 145)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 146
'...........................................................................................
'aggiunge campi per gestione informazioni disabili nella tabella Luoghi
'...........................................................................................
sql = "ALTER TABLE tb_attivita ADD att_inSospeso BIT NULL"
CALL DB.Execute(sql, 146)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 147
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
	  "UPDATE tb_siti SET sito_nome='NEXT-banner [gestione banners pubblicitari]' WHERE id_sito=" & NEXTBANNER & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-club [gestione associati]' WHERE id_sito=" & NEXTCLUB & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-booking [gestione prenotazioni]' WHERE id_sito=" & NEXTBOOKING & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-guestbook [gestione guestbook]' WHERE id_sito=" & NEXTGUESTBOOK & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-contract [gestione bandi ed appalti] ' WHERE id_sito=" & NEXTCONTRACT & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-f.a.q. [gestione frequently asked questions]' WHERE id_sito=" & NEXTFAQ & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-team [gestione organigramma aziendale]' WHERE id_sito=" & NEXTTEAM & ";" & vbCrLf & _
	  "UPDATE tb_siti SET sito_nome='NEXT-booking portal [gestione portale di prenotazione]' WHERE id_sito=" & NEXTBOOKINGPORTALE
CALL DB.Execute(sql, 147)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 148
'...........................................................................................
'aggiunge campi per gestione informazioni disabili nella tabella Luoghi
'...........................................................................................
sql = "ALTER TABLE tb_pratiche ADD " & vbCrLf & _
	  "		pra_mod_data SMALLDATETIME NULL, " & vbCrLf & _
	  "		pra_mod_utente INTEGER NULL; " & vbCrLf & _
	  "ALTER TABLE tb_documenti ADD " & vbCrLf & _
	  "		doc_mod_data SMALLDATETIME NULL, " & vbCrLf & _
	  "		doc_mod_utente INTEGER NULL"
CALL DB.Execute(sql, 148)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 149
'...........................................................................................
'aggiornamento fantasma per la creazione delle directory temporanee per ogni utente
'...........................................................................................
sql = "SELECT * FROM tb_admin"
CALL DB.Execute(sql, 149)
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
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adAsyncFetch
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
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 150
'...........................................................................................
'aggiunge tabelle per la gestione dei files
'...........................................................................................
sql = " CREATE TABLE dbo.tb_Files (" & _
	  "		F_id INT IDENTITY (1, 1) NOT NULL, " & _
	  "		F_original_name varchar(250) NULL, " & _
	  "		F_encoded_name varchar(250) NULL, " & _
	  "		F_size INT NULL, " & _
	  "		F_Data SMALLDATETIME NULL, " & _
	  "		F_base_path varchar(250) NULL, " & _
	  "		F_allegato BIT NULL " &_
	  "		);" & _
	  " CREATE TABLE dbo.rel_documenti_files(" & _
	  "		rel_id INT IDENTITY (1, 1) NOT NULL, " & _
	  "		rel_documento_id INT NOT NULL, " & _
	  " 	rel_files_id INT NOT NULL " & _
	  "		);" & _
	  "	ALTER TABLE tb_Files ADD CONSTRAINT PK_tb_Files PRIMARY KEY (F_id); " &_
	  "	ALTER TABLE rel_documenti_files ADD CONSTRAINT PK_rel_documenti_files PRIMARY KEY (rel_id); " &_
	  " ALTER TABLE rel_documenti_files ADD CONSTRAINT FK_rel_documenti_files__tb_files " + _
   	  " 	FOREIGN KEY (rel_files_id) REFERENCES tb_files (F_id) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; " & _
	  " ALTER TABLE rel_documenti_files ADD CONSTRAINT FK_rel_documenti_files__tb_documenti " + _
	  "		FOREIGN KEY (rel_documento_id) REFERENCES tb_documenti (doc_id) " + _
	  "		ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 150)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 151
'...........................................................................................
'aggiunge tabelle per la gestione dei files
'...........................................................................................
sql = " CREATE PROCEDURE dbo.Delete_PuntiStrat (@IDPunto int) AS " & vbCrLF & _
	  "		DELETE FROM RelazPS WHERE ID_PuntoStr = @IDPunto " & vbCrLF & _
	  "		DELETE FROM PuntiStrat WHERE IDPunto = @IDPunto "
CALL DB.Execute(sql, 151)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 152
'...........................................................................................
'aggiunge campi a tabella files e crea directory docs
'...........................................................................................
sql = " ALTER TABLE tb_files ADD " + _
	  "		F_original_path varchar(250) NULL, " & _
	  " 	F_encoded_path varchar(250) NULL, " & _
	  "		F_LastUpdate SMALLDATETIME NULL; " & _
	  " ALTER TABLE tb_files DROP COLUMN F_base_path "
CALL DB.Execute(sql, 152)
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
'AGGIORNAMENTO 153
'...........................................................................................
'aggiunge campo per tracciatura utente che chiude l'attivita'
'...........................................................................................
sql = " ALTER TABLE tb_attivita ADD " + _
	  "		att_utente_chiusura INT NULL "
CALL DB.Execute(sql, 153)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 154
'...........................................................................................
'elimina vecchia gestione documenti
'...........................................................................................
sql = " ALTER TABLE tb_documenti DROP COLUMN doc_path; " + _
	  " DELETE FROM tb_files; " & _
	  " DELETE FROM tb_documenti; "
CALL DB.Execute(sql, 154)
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
'AGGIORNAMENTO 155
'...........................................................................................
'svuota tabelle documenti e files e corregge problema su relazione tb_allegati e tb_documenti
'...........................................................................................
sql = " DELETE FROM tb_files; " & _
	  " DELETE FROM tb_documenti; " & _
	  " DELETE FROM tb_allegati; " & _
	  " ALTER TABLE tb_allegati ADD CONSTRAINT FK_tb_allegati__tb_documenti " + _
   	  " FOREIGN KEY (all_documento_id) REFERENCES tb_documenti (doc_id) " + _
	  " ON UPDATE CASCADE ON DELETE CASCADE; "
CALL DB.Execute(sql, 155)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 156
'...........................................................................................
'aumenta lunghezza campo nome del sito
'...........................................................................................
sql = " ALTER TABLE tb_siti ALTER COLUMN sito_nome nvarchar(250) NULL"
CALL DB.Execute(sql, 156)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 157
'...........................................................................................
'aumenta dimensione campo email su log di spedizione
'...........................................................................................
sql = "ALTER TABLE log_cnt_email ALTER COLUMN log_email nvarchar(250) "
CALL DB.Execute(sql, 157)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 158
'...........................................................................................
'aggiornamento del sistema per portare tutto a stand alone
'...........................................................................................
sql = " UPDATE tb_siti SET sito_dir='NEXTPassport' WHERE id_sito=" & NEXTPASSPORT & "; " + _
	  " UPDATE tb_siti SET sito_dir='NEXTWeb' WHERE id_sito=" & NEXTWEB & "; " + _
	  " UPDATE tb_siti SET sito_dir='AptCircolari' WHERE id_sito=" & NEXTMEMO & "; " + _
	  " UPDATE tb_siti SET sito_dir='NEXTcom' WHERE id_sito=" & NEXTCOM & "; " + _
	  " UPDATE tb_siti SET sito_dir='AptControlloQ' WHERE id_sito=" & APT_QUALITA & "; " + _
	  " UPDATE tb_siti SET sito_dir='AptAdmin' WHERE id_sito=" & APT_ADMIN
CALL DB.Execute(sql, 158)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 159
'...........................................................................................
'aggiunge campo su contatti per indicare se il contatto interno &egrave; in un'altra sede
'...........................................................................................
sql = " ALTER TABLE tb_indirizzario ADD " + _
	  "		altra_sede BIT NULL "
CALL DB.Execute(sql, 159)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 160
'...........................................................................................
'aggiunge campi descrizione per stru_ric
'...........................................................................................
sql = " ALTER TABLE stru_ric ADD " + vbCrlf + _
	  "		descr_it ntext NULL, " + vbCrlf + _
	  "		descr_en ntext NULL, " + vbCrlf + _
	  "		descr_fr ntext NULL, " + vbCrlf + _
	  "		descr_es ntext NULL, " + vbCrlf + _
	  "		descr_de ntext NULL"
CALL DB.Execute(sql, 160)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 161
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
	  "UPDATE tb_siti SET sito_dir='NEXTweb' WHERE id_sito=25 ;" + vbCrLf + _
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
CALL DB.Execute(sql, 161)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 162
'...........................................................................................
'corregge nome applicatiivi nextCom e nextDoc+
'...........................................................................................
sql = "UPDATE tb_siti SET sito_nome='" + IIF(Application("NextCrm"), "NEXT-doc+ [comunicazioni &amp;; documenti]", "NEXT-com [gestione comunicazioni]") + "' WHERE id_sito=3; "
CALL DB.Execute(sql, 162)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 163
'...........................................................................................
'corregge nome applicativo next-web 4
'...........................................................................................
sql = "UPDATE tb_siti SET sito_nome='NEXT-web 4.0 [gestione grafica e contenuti]' WHERE id_sito=25"
CALL DB.Execute(sql, 163)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 164
'...........................................................................................
'corregge problema di report per applicativo "presenze"
'...........................................................................................
sql = "ALTER PROCEDURE dbo.GET_DIP_Stringa_Richieste ( " + vbCrLf + _
	  "		@id_admin int, " + vbCrLf + _
	  "		@data nvarchar(30), " + vbCrLf + _
	  "		@stringa nvarchar(500) OUTPUT  " + vbCrLf + _
	  "	) AS " + vbCrLf + _
	  "		DECLARE @val int, @nome_Serv nvarchar(50), @nome_tipo nvarchar(50), @data_inizio nvarchar(20), @data_fine nvarchar(20) " + vbCrLf + _
	  "		DECLARE rs CURSOR FOR " + vbCrLf + _
	  "		SELECT tb_richieste.valore_richiesta, tb_servizi.nome_servizio, tb_tipiservizi.nometiposerv,  " + vbCrLf + _
	  "		CONVERT(nvarchar(20),tb_richieste.data_inizio, 105), CONVERT(nvarchar(20),tb_richieste.data_fine, 105) " + vbCrLf + _
	  "		FROM tb_richieste INNER JOIN  " + vbCrLf + _
	  "		(tb_servizi INNER JOIN tb_tipiServizi  " + vbCrLf + _
	  "		ON tb_servizi.tipo_servizio=tb_tipiServizi.idTiposerv)  " + vbCrLf + _
	  "		ON tb_richieste.tipo_richiesta=tb_servizi.id_servizio  " + vbCrLf + _
	  "		WHERE id_dipendente = @id_admin AND (tb_richieste.tipo_messaggio='RI' OR tb_richieste.tipo_messaggio='N')  " + vbCrLf + _
	  "		AND ISNULL(tb_richieste.valore_richiesta, 0)<>0  " + vbCrLf + _
	  "		AND (tb_richieste.data_inizio <= CONVERT(DATETIME, @data , 102) AND tb_richieste.data_fine >= CONVERT(DATETIME, @data , 102)) " + vbCrLf + _
	  "		OPEN rs " + vbCrLf + _
	  "		SET @stringa = '' " + vbCrLf + _
	  "		FETCH NEXT FROM rs INTO @val, @nome_Serv, @nome_tipo, @data_inizio, @data_fine " + vbCrLf + _
	  "		WHILE (@@FETCH_STATUS=0) " + vbCrLf + _
	  "			BEGIN " + vbCrLf + _
	  "				SET @stringa = @stringa + '(' + str(@val,3,0) + ' ) '  " + vbCrLf + _
	  "				IF (@val<0) BEGIN " + vbCrLf + _
	  "					SET @stringa = @stringa + 'Utilizzo ' " + vbCrLf + _
	  "				END " + vbCrLf + _
	  "				SET @stringa = @stringa + @nome_tipo + ' di ' + @nome_Serv + ' dal: ' + @data_inizio + ' al: ' + @data_fine + ';; ' " + vbCrLf + _
	  "				FETCH NEXT FROM rs INTO @val, @nome_Serv, @nome_tipo, @data_inizio, @data_fine " + vbCrLf + _
	  "			END " + vbCrLf + _
	  "		CLOSE rs " + vbCrLf + _
	  "		DEALLOCATE rs " + vbCrLf + _
	  "		return; "
CALL DB.Execute(sql, 164)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 165
'...........................................................................................
'aggiunge campi per descrittori: ordine, flag principale
'...........................................................................................
sql = " ALTER TABLE tb_descrittori ADD " + _
	  "		descr_ordine INT NULL, " + _
	  " 	descr_principale BIT NULL; " + _
	  " UPDATE tb_descrittori SET descr_ordine = 0 "
CALL DB.Execute(sql, 165)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 166
'...........................................................................................
'aggiorna campi parametri applicativi
'...........................................................................................
sql = " ALTER TABLE tb_siti_parametri ALTER COLUMN " + _
	  "		par_key nvarchar(250) NULL; " + _
	  " ALTER TABLE tb_siti_parametri ALTER COLUMN " + _
	  "		par_value nvarchar(250) NULL "
CALL DB.Execute(sql, 166)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 167
'...........................................................................................
'rimuove applicativo NEXT-MEMO installando al suo posto l'applicativo APT circolari
'gemerando una nuova tabella per la gestione delle circolari
'...........................................................................................
sql = " INSERT INTO tb_siti (id_sito, sito_nome, sito_dir, sito_p1, sito_p2, sito_p3, sito_amministrazione) " + _
	  " SELECT " & APT_CIRCOLARI & ", sito_nome, 'AptCircolari', sito_p1, sito_p2, sito_p3, 1 " + _
	  " FROM tb_siti WHERE id_sito = " & NEXTMEMO & " ; " + _
	  " UPDATE log_admin SET log_sito_id=" & APT_CIRCOLARI & " WHERE log_sito_id=" & NEXTMEMO & "; " + _
	  " UPDATE rel_admin_sito SET sito_id = " & APT_CIRCOLARI & " WHERE sito_id=" & NEXTMEMO & "; " + _
	  " UPDATE rel_utenti_sito SET rel_sito_id = " & APT_CIRCOLARI & " WHERE rel_sito_id=" & NEXTMEMO & "; " + _
	  " UPDATE log_utenti SET log_sito_id=" & APT_CIRCOLARI & " WHERE log_sito_id=" & NEXTMEMO & "; " + _
	  " UPDATE tb_siti_parametri SET par_sito_id=" & APT_CIRCOLARI & " WHERE par_sito_id=" & NEXTMEMO & "; " + _
	  " DELETE FROM tb_siti WHERE id_sito =" & NEXTMEMO & "; " + _
	  " CREATE TABLE dbo.Tb_APT_Circolari ( " + _
	  " 	CI_ID int IDENTITY (1, 1) NOT NULL , " + _
	  "		CI_Numero nvarchar (50) NULL , " + _
	  " 	CI_Titolo nvarchar (250) NULL , " + _
	  "		CI_Estratto ntext NULL , " + _
	  "		CI_Pubblicazione datetime NULL , " + _
	  "		CI_Scadenza datetime NULL , " + _
	  "		CI_File nvarchar (100) NULL , " + _
	  "		CI_Visibile bit NULL , " + _
	  "		CI_protetto bit NOT NULL , " + _
	  "		CONSTRAINT PK_tb_APT_circolari PRIMARY KEY NONCLUSTERED (CI_ID) " + _
	  "	) ; " + _
	  " INSERT INTO Tb_APT_Circolari (CI_Numero, CI_Titolo, CI_Estratto, CI_Pubblicazione, CI_Scadenza, CI_File, CI_Visibile, CI_protetto) " + _
	  " 					SELECT 	  CI_Numero, CI_Titolo, CI_Estratto, CI_Pubblicazione, CI_Scadenza, CI_File, CI_Visibile, CI_protetto " + _
	  "						FROM tb_circolari ; " + _
	  " CREATE TABLE dbo.log_APT_circolari ( " + _
	  "		log_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		log_ut_id int NOT NULL , " + _
	  "		log_dip_id int NOT NULL , " + _
	  "		log_ci_id int NOT NULL , " + _
	  "		log_data datetime NULL , " + _
	  "		CONSTRAINT PK_log_APT_circolari PRIMARY KEY NONCLUSTERED (log_id), " + _
	  "		CONSTRAINT FK_log_APT_circolari__tb_APT_Circolari FOREIGN KEY (log_ci_id) " + _
	  "		REFERENCES Tb_APT_Circolari (CI_ID) ON DELETE CASCADE  ON UPDATE CASCADE " + _
	  "	) ; " + _
	  " INSERT INTO log_APT_circolari (log_ut_id, log_dip_id, log_ci_id, log_data ) " + _
	  "							SELECT log_ut_id, log_dip_id, log_ci_id, log_data " + _
	  "							FROM log_APT_circolari ; " + _
	  " ALTER TABLE dbo.log_circolari DROP CONSTRAINT FK_log_circolari__tb_Circolari; " + _
	  DropObject(DB.objConn, "tb_circolari", "TABLE") + _
	  DropObject(DB.objConn, "log_circolari", "TABLE") + _
	  " UPDATE tb_siti SET sito_dir='../AptCircolari' WHERE id_sito=" & APT_CIRCOLARI & "; " + _
	  " UPDATE tb_siti SET sito_dir='../AptControlloQ' WHERE id_sito=" & APT_QUALITA & "; "
CALL DB.Execute(sql, 167)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 168
'...........................................................................................
'controlla e ripulisce directory vecchie e residue
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 168)
if DB.last_update_executed then
	CALL Aggiornamento__FRAMEWORK_CORE__pulizia_directory(conn, rs)
end if
'...........................................................................................
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 169
'...........................................................................................
'installa applicazioni per portare il sistema a FRAMEWORK CORE
'esegue aggiornamento 1
'...........................................................................................
sql = Install__FRAMEWORK_CORE__NEXTLINK(conn) + _
	  Install__FRAMEWORK_CORE__NEXTNEWS(conn) + _
	  Install__FRAMEWORK_CORE__NEXTGALLERY(conn) + _
	  Install__FRAMEWORK_CORE__NEXTFAQ(conn) + _
	  Install__FRAMEWORK_CORE__NEXTTEAM(conn) + _
	  Aggiornamento__FRAMEWORK_CORE__1(conn)
CALL DB.Execute(sql, 169)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 170
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__5(conn)
CALL DB.Execute(sql, 170)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 171
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__6(conn)
CALL DB.Execute(sql, 171)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 172
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__7(conn)
CALL DB.Execute(sql, 172)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 173
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__8(conn)
CALL DB.Execute(sql, 173)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 174
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__9(conn)
CALL DB.Execute(sql, 174)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 175
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__10(conn)
CALL DB.Execute(sql, 175)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 176
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__11(conn)
CALL DB.Execute(sql, 176)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 177
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__12(DB, rs, 177)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 178
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__13(conn)
CALL DB.Execute(sql, 178)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 179
'...........................................................................................
sql = rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
CALL DB.Execute(sql, 179)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 180
'...........................................................................................
sql = Install__FRAMEWORK_CORE__NEXTWEB5(conn)
CALL DB.Execute(sql, 180)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 181
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__14(conn)
CALL DB.Execute(sql, 181)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 182
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__15(conn)
CALL DB.Execute(sql, 182)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 183
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__16(conn)
CALL DB.Execute(sql, 183)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 184
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__17(conn)
CALL DB.Execute(sql, 184)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 185
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__18(conn)
CALL DB.Execute(sql, 185)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 186
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__19(conn)
CALL DB.Execute(sql, 186)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 187
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__20(conn)
CALL DB.Execute(sql, 187)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 188
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__21(conn)
CALL DB.Execute(sql, 188)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 189
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__22(conn)
CALL DB.Execute(sql, 189)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 190
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__23(conn)
CALL DB.Execute(sql, 190)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 191
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__24(conn)
CALL DB.Execute(sql, 191)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 192
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__25(conn)
CALL DB.Execute(sql, 192)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 193
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__26(conn) + _
	  Aggiornamento__FRAMEWORK_CORE__27(conn)
CALL DB.Execute(sql, 193)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 194
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__28(conn)
CALL DB.Execute(sql, 194)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 195
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__29(conn)
CALL DB.Execute(sql, 195)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 196
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__30(conn)
CALL DB.Execute(sql, 196)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 197
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__31(conn)
CALL DB.Execute(sql, 197)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 198
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__32(conn)
CALL DB.Execute(sql, 198)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 199
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__33(conn)
CALL DB.Execute(sql, 199)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 200
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__34(conn)
CALL DB.Execute(sql, 200)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 201
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__35(conn)
CALL DB.Execute(sql, 201)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 202
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__36(conn)
CALL DB.Execute(sql, 202)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 203
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__37(conn)
CALL DB.Execute(sql, 203)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 204
'...........................................................................................
sql = AggiornamentoSpeciale__FRAMEWORK_CORE__38(DB, rs, 204)
CALL DB.Execute(sql, 204)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 205
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__39(conn)
CALL DB.Execute(sql, 205)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 206
'...........................................................................................
'corregge problema rigenerazione vista
'...........................................................................................
sql = " DROP VIEW dbo.viewLocalieServizi; " + _
      " CREATE VIEW dbo.viewLocalieServizi AS " + vbCrLf + _
      "     SELECT dbo.LocalieServizi.*," + vbCrLf + _
      "            dbo.Tipi_LS.tipo_nome_it, dbo.Tipi_LS.tipo_nome_eng, dbo.Tipi_LS.tipo_nome_fra, dbo.Tipi_LS.tipo_nome_ted, dbo.Tipi_LS.tipo_nome_spa " + vbCrLF + _
      "     FROM dbo.LocalieServizi INNER JOIN dbo.Tipi_LS ON dbo.LocalieServizi.Tipo = dbo.Tipi_LS.id_tipoutil "
CALL DB.Execute(sql, 206)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(206)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 207
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__40(conn)
CALL DB.Execute(sql, 207)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 208
'...........................................................................................
'cambia campi dei prezzi delle spiagge da interi a valori con la virgola
'...........................................................................................
sql = " ALTER TABLE Spiagge ALTER COLUMN ombr_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN ombr_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN ombr_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN ombr_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN ombr_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN ombr_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN lett_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN spo_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN camer_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN minicap_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf1_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf2_stag_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_giorn_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_giorn_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_mens_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_mens_max real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_stag_min real NULL; " + _
      " ALTER TABLE Spiagge ALTER COLUMN capf3_stag_max real NULL; "
CALL DB.Execute(sql, 208)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 209
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__41(conn)
CALL DB.Execute(sql, 209)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 210
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__42(conn)
CALL DB.Execute(sql, 210)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 211
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__43(conn)
CALL DB.Execute(sql, 211)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 212
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__44(conn)
CALL DB.Execute(sql, 212)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTI SALTATI PER REVISIONE SCRIPT DI UPDATE
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 213)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 214)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 215)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 216)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 217
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__48(conn)
CALL DB.Execute(sql, 217)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 218
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__49(conn)
CALL DB.Execute(sql, 218)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 219
'...........................................................................................
'AGGIORNAMENTI SALTATI PER REVISIONE SCRIPT DI UPDATE
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 219)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 220
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__50(conn)
CALL DB.Execute(sql, 220)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 221
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__51(conn)
CALL DB.Execute(sql, 221)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 222
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__52(conn)
CALL DB.Execute(sql, 222)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTI SALTATI PER REVISIONE SCRIPT DI UPDATE
'...........................................................................................
sql = "SELECT * FROM aa_versione"
CALL DB.Execute(sql, 223)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 224
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__53(conn)
CALL DB.Execute(sql, 224)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 225
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__54(conn)
CALL DB.Execute(sql, 225)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTI SALTATI PER REVISIONE SCRIPT DI UPDATE
'...........................................................................................
sql = "SELECT * FROM aa_versione"
CALL DB.Execute(sql, 226)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 227
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__56(conn)
CALL DB.Execute(sql, 227)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 228
'...........................................................................................
'rinomina struttura dati del carnet apt per permettere la creazione del carnet del framework
'...........................................................................................
sql = " DELETE FROM tb_appunti; " + _
      " ALTER TABLE tb_appunti DROP CONSTRAINT FK__tb_appunt__carne__3A978D17 ; " + _
      DropObject(conn, "tb_carnet", "TABLE") + _
      " CREATE TABLE dbo.tb_carnet_apt ( " + _
      "     id_carnet int IDENTITY(1, 1) NOT NULL , " + _
      "     session_carnet nvarchar(50) NULL , " + _
      "     data_carnet smalldatetime NULL , " + _
      "     sito_carnet int NULL , " + _
      "     PRIMARY KEY CLUSTERED ( id_carnet ) " + _
      " ) ; " + _
      " ALTER TABLE dbo.tb_appunti ADD CONSTRAINT FK__tb_appunti__tb_carnet_apt " + _
      "     FOREIGN KEY (carnet_appunto) REFERENCES tb_carnet_apt(id_carnet) "
CALL DB.Execute(sql, 228)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 229
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__45(conn)
CALL DB.Execute(sql, 229)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 230
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__46(conn)
CALL DB.Execute(sql, 230)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 231
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__47(conn)
CALL DB.Execute(sql, 231)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 232
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__55(conn)
CALL DB.Execute(sql, 232)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 233
'...........................................................................................
'cambia campo "sito_carnet" da codice numerico a codice di caratteri
'...........................................................................................
sql = " DELETE FROM tb_appunti; " + _
      " DELETE FROM tb_carnet_apt; " + _
      " ALTER TABLE tb_carnet_apt ALTER COLUMN sito_carnet nvarchar(50) NULL; " + _
      " ALTER TABLE tb_appunti DROP CONSTRAINT FK__tb_appunti__tb_carnet_apt; " + _
      SQL_AddForeignKey(conn, "tb_appunti", "carnet_appunto", "tb_carnet_apt", "id_carnet", true, "")
CALL DB.Execute(sql, 233)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 234
'...........................................................................................
'aggiornamento per la pulizia delle directory non utilizzate
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 234)
if DB.last_update_executed then
	CALL Aggiornamento_234_PulituraDirectory(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_234_PulituraDirectory(conn, rs)
    dim UploadPath
    
    set fso = Server.CreateObject("Scripting.FileSystemObject")
    UploadPath = Application("IMAGE_PATH")
    
    'rimuove cartelle inutili su upload:
    CALL FolderRemove(fso, UploadPath + "XML_cache", false)

    CALL ClearTempDir(fso)
    
    set fso = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 235
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__57(conn)
CALL DB.Execute(sql, 235)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 236
'...........................................................................................
'aggiorna dati dell'applicazione per collegarla al next-info della provincia
'...........................................................................................
sql = " UPDATE tb_siti SET " + _
	  "		sito_dir='http://www.turismo.provincia.venezia.it/amministrazione/nextInfo', " + _
	  "		sito_nome='APT - Amministrazione dati informativi Provincia di Venezia' " + _
	  " WHERE id_sito=" & APT_ADMIN
CALL DB.Execute(sql, 236)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 237
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__58(conn)
CALL DB.Execute(sql, 237)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 238
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__59(conn)
CALL DB.Execute(sql, 238)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 239
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__60(conn)
CALL DB.Execute(sql, 239)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 240
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__61(conn)
CALL DB.Execute(sql, 240)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(240)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 241
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__62(conn)
CALL DB.Execute(sql, 241)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 242
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__63(conn)
CALL DB.Execute(sql, 242)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 243
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__64(conn)
CALL DB.Execute(sql, 243)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 244
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__65(conn)
CALL DB.Execute(sql, 244)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 245
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__66(conn)
CALL DB.Execute(sql, 245)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 246
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__67(conn)
CALL DB.Execute(sql, 246)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 247
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__68(conn)
CALL DB.Execute(sql, 247)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 248
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__69(conn)
CALL DB.Execute(sql, 248)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 249
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__70(conn)
CALL DB.Execute(sql, 249)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 250
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__71(conn)
CALL DB.Execute(sql, 250)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 251
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__72(DB, 251)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 252
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__73(conn)
CALL DB.Execute(sql, 252)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 253
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__74(conn)
CALL DB.Execute(sql, 253)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 254
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__75(conn)
CALL DB.Execute(sql, 254)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(254)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 255
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__76(conn)
CALL DB.Execute(sql, 255)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 256
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__77(conn)
CALL DB.Execute(sql, 256)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 257
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__78(conn)
CALL DB.Execute(sql, 257)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 258
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__79(conn)
CALL DB.Execute(sql, 258)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 259
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__80(conn)
CALL DB.Execute(sql, 259)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 260
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__81(conn)
CALL DB.Execute(sql, 260)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(260)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 261
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__82(conn)
CALL DB.Execute(sql, 261)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 262
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__83(conn)
CALL DB.Execute(sql, 262)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 263
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__84(conn)
CALL DB.Execute(sql, 263)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 264
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__85(conn)
CALL DB.Execute(sql, 264)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 265
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__86(conn)
CALL DB.Execute(sql, 265)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 266
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__87(DB, 266)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 267
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__88(conn)
CALL DB.Execute(sql, 267)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 268
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__89(conn)
CALL DB.Execute(sql, 268)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(268)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 269
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__90(DB, rs, 269)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 270
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__91(conn)
CALL DB.Execute(sql, 270)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(270)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 271
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__92(conn)
CALL DB.Execute(sql, 271)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 272
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__93(conn)
CALL DB.Execute(sql, 272)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 273
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__94(conn)
CALL DB.Execute(sql, 273)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 274
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__95(conn)
CALL DB.Execute(sql, 274)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.RebuildIndex_RefreshContents("tb_webs", "id_webs")
'*******************************************************************************************

'*******************************************************************************************
CALL DB.RebuildIndex_RefreshContents("tb_paginesito", "id_pagineSito")
'*******************************************************************************************

'*******************************************************************************************
CALL DB.RebuildIndex_RefreshContents("tb_contents_index", "idx_livello")
'*******************************************************************************************

'*******************************************************************************************
CALL DB.RebuildIndex_OperazioniRicorsive()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 275
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__96(conn)
CALL DB.Execute(sql, 275)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 276
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__97(conn)
CALL DB.Execute(sql, 276)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 277
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__98(conn)
CALL DB.Execute(sql, 277)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 278
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__99(conn)
CALL DB.Execute(sql, 278)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 279
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__100(conn)
CALL DB.Execute(sql, 279)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(279)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 280
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__101(conn)
CALL DB.Execute(sql, 280)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 281
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__102(conn)
CALL DB.Execute(sql, 281)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 282
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__103(conn)
CALL DB.Execute(sql, 282)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(282)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 283
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__104(conn)
CALL DB.Execute(sql, 283)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 284
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__105(conn)
CALL DB.Execute(sql, 284)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.RebuildIndex_RefreshContents("tb_webs", "id_webs")
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 285
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__106(conn)
CALL DB.Execute(sql, 285)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 286
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__107(conn)
CALL DB.Execute(sql, 286)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 287
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__108(conn)
CALL DB.Execute(sql, 287)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 288
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__109(conn)
CALL DB.Execute(sql, 288)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 289
'...........................................................................................
'elimina vecchia struttura dati dei portali APT
'...........................................................................................
sql = DropObject(conn, "Delete_Admin", "PROCEDURE") + _
	  DropObject(conn, "Delete_ALL_eventi", "PROCEDURE") + _
	  DropObject(conn, "delete_ALL_Itinerari", "PROCEDURE") + _
	  DropObject(conn, "delete_ALL_LS", "PROCEDURE") + _
	  DropObject(conn, "delete_ALL_luoghi", "PROCEDURE") + _
	  DropObject(conn, "delete_ALL_NotizieUtili", "PROCEDURE") + _
	  DropObject(conn, "delete_ALL_strutture", "PROCEDURE") + _
	  DropObject(conn, "Delete_Caratt_TipiRic", "PROCEDURE") + _
	  DropObject(conn, "Delete_Categoria", "PROCEDURE") + _
	  DropObject(conn, "Delete_CategorieEventi", "PROCEDURE") + _
	  DropObject(conn, "Delete_Eventi", "PROCEDURE") + _
	  DropObject(conn, "Delete_EventiSpeciali", "PROCEDURE") + _
	  DropObject(conn, "DELETE_LS", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Luoghi", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Not_Util", "PROCEDURE") + _
	  DropObject(conn, "Delete_PuntiStrat", "PROCEDURE") + _
	  DropObject(conn, "Delete_Servizi_TipiRic", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Spiagge", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Stru_Ric", "PROCEDURE") + _
	  DropObject(conn, "Delete_SubTipo_LS", "PROCEDURE") + _
	  DropObject(conn, "Delete_SUbTipo_Not", "PROCEDURE") + _
	  DropObject(conn, "Delete_SubZona", "PROCEDURE") + _
	  DropObject(conn, "Delete_tb_Admin", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Tipi_ric", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Tipo_LS", "PROCEDURE") + _
	  DropObject(conn, "DELETE_Tipo_Not_Util", "PROCEDURE") + _
	  DropObject(conn, "DELETE_TipoLuoghi", "PROCEDURE") + _
	  DropObject(conn, "Delete_Zona", "PROCEDURE") + _
	  DropObject(conn, "viewAllNotUtili", "VIEW") + _
	  DropObject(conn, "viewDotazioni", "VIEW") + _
	  DropObject(conn, "viewEventi", "VIEW") + _
	  DropObject(conn, "viewEventiCompleto", "VIEW") + _
	  DropObject(conn, "viewEventiPerCategorie", "VIEW") + _
	  DropObject(conn, "viewEventiPerLuogo", "VIEW") + _
	  DropObject(conn, "viewLocalieServizi", "VIEW") + _
	  DropObject(conn, "viewNotUtili", "VIEW") + _
	  DropObject(conn, "viewServizi", "VIEW") + _
	  DropObject(conn, "viewTipiRic", "VIEW") + _
	  " ALTER TABLE doveAccade DROP CONSTRAINT FK__doveAccad__id_ev__3493CFA7 ; " + _
	  " ALTER TABLE doveAccade DROP CONSTRAINT FK__doveAccad__id_lu__339FAB6E ; " + _
	  " ALTER TABLE Eventi DROP CONSTRAINT FK__Eventi__id_categ__22751F6C ; " + _
	  " ALTER TABLE imag_ev DROP CONSTRAINT FK__imag_ev__id_ev__25518C17 ; " + _
	  " ALTER TABLE imag_ls DROP CONSTRAINT FK_imag_ls_LocalieServizi ; " + _
	  " ALTER TABLE imag_lu DROP CONSTRAINT FK__imag_lu__Id_luog__30C33EC3 ; " + _
	  " ALTER TABLE Luoghi DROP CONSTRAINT FK__Luoghi__id_tipo__2CF2ADDF ; " + _
	  " ALTER TABLE Luoghi DROP CONSTRAINT FK__Luoghi__Zona__2DE6D218 ; " + _
	  " ALTER TABLE Not_util DROP CONSTRAINT FK_Not_util_Tipi_notutil ; " + _
	  " ALTER TABLE Not_util DROP CONSTRAINT FK_Not_util_Zone_StruRic ; " + _
	  " ALTER TABLE PeriodiEventi DROP CONSTRAINT FK__PeriodiEv__Id_ev__282DF8C2 ; " + _
	  " ALTER TABLE rel_ric_caratt DROP CONSTRAINT FK__rel_ric_c__rel_i__4CC05EF3 ; " + _
	  " ALTER TABLE rel_ric_caratt DROP CONSTRAINT FK__rel_ric_c__rel_i__1AD3FDA4 ; " + _
	  " ALTER TABLE rel_Ric_Servizi DROP CONSTRAINT FK__rel_Ric_S__rel_r__53633AE1 ; " + _
	  " ALTER TABLE rel_Ric_Servizi DROP CONSTRAINT FK__rel_Ric_S__rel_S__123EB7A3 ; " + _
	  " ALTER TABLE rel_Serv_TipiRic DROP CONSTRAINT FK__rel_Serv___rel_t__35D2D7FA ; " + _
	  " ALTER TABLE rel_Serv_TipiRic DROP CONSTRAINT PK__rel_Serv_TipiRic__7BA63665 ; " + _
	  " ALTER TABLE rel_sottoTipi_LS DROP CONSTRAINT FK_rel_sottoTipi_LS_LocalieServizi ; " + _
	  " ALTER TABLE rel_sottoTipi_LS DROP CONSTRAINT FK_rel_sottoTipi_LS_SottoTipi_LocServ ; " + _
	  " ALTER TABLE rel_sottoTipi_NotUtil DROP CONSTRAINT FK__rel_sotto__rel_s__5B045CA9 ; " + _
	  " ALTER TABLE rel_sottoTipi_NotUtil DROP CONSTRAINT FK__rel_sotto__rel_s__58D1301D ; " + _
	  " ALTER TABLE rel_tipiStr_caratt DROP CONSTRAINT FK__rel_tipiS__rel_i__37BB206C ; " + _
	  " ALTER TABLE rel_tipiStr_caratt DROP CONSTRAINT PK__rel_tipiStr_cara__7C9A5A9E ; " + _
	  " ALTER TABLE RelazPS DROP CONSTRAINT FK__RelazPS__ID_Punt__4BC21919 ; " + _
	  " ALTER TABLE RelazPS DROP CONSTRAINT FK__RelazPS__ID_Stru__08B54D69 ; " + _
	  " ALTER TABLE SottoTipi_LocServ DROP CONSTRAINT FK_SottoTipi_LocServ_Tipi_LS ; " + _
	  " ALTER TABLE SottoTipi_Notutil DROP CONSTRAINT PK__SottoTipi_Notuti__77D5A581 ; " + _
	  " ALTER TABLE Stru_ric DROP CONSTRAINT FK__Stru_ric__Zona__02FC7413 ; " + _
	  " ALTER TABLE SubZone DROP CONSTRAINT FK_SubZone_Zone_StruRic ; " + _
	  DropObject(conn, "Caratt_TipiRic", "TABLE") + _
	  DropObject(conn, "CategorieEventi", "TABLE") + _
	  DropObject(conn, "doveAccade", "TABLE") + _
	  DropObject(conn, "Eventi", "TABLE") + _
	  DropObject(conn, "eventiSpeciali", "TABLE") + _
	  DropObject(conn, "imag_ev", "TABLE") + _
	  DropObject(conn, "imag_ls", "TABLE") + _
	  DropObject(conn, "imag_lu", "TABLE") + _
	  DropObject(conn, "LocalieServizi", "TABLE") + _
	  DropObject(conn, "Luoghi", "TABLE") + _
	  DropObject(conn, "Not_util", "TABLE") + _
	  DropObject(conn, "PeriodiEventi", "TABLE") + _
	  DropObject(conn, "PuntiStrat", "TABLE") + _
	  DropObject(conn, "rel_ric_caratt", "TABLE") + _
	  DropObject(conn, "rel_Ric_Servizi", "TABLE") + _
	  DropObject(conn, "rel_Serv_TipiRic", "TABLE") + _
	  DropObject(conn, "rel_sottoTipi_LS", "TABLE") + _
	  DropObject(conn, "rel_sottoTipi_NotUtil", "TABLE") + _
	  DropObject(conn, "rel_tipiStr_caratt", "TABLE") + _
	  DropObject(conn, "RelazPS", "TABLE") + _
	  DropObject(conn, "Servizi_TipiRic", "TABLE") + _
	  DropObject(conn, "SottoTipi_LocServ", "TABLE") + _
	  DropObject(conn, "SottoTipi_Notutil", "TABLE") + _
	  DropObject(conn, "Spiagge", "TABLE") + _
	  DropObject(conn, "Stru_ric", "TABLE") + _
	  DropObject(conn, "Stru_Ric_WS_Config", "TABLE") + _
	  DropObject(conn, "SubZone", "TABLE") + _
	  DropObject(conn, "tb_quickmail", "TABLE") + _
	  DropObject(conn, "Tipi_LS", "TABLE") + _
	  DropObject(conn, "Tipi_notutil", "TABLE") + _
	  DropObject(conn, "Tipi_ric", "TABLE") + _
	  DropObject(conn, "TipoLuoghi", "TABLE") + _
	  DropObject(conn, "Zone_StruRic", "TABLE")
CALL DB.Execute(sql, 289)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 290
'...........................................................................................
'non eseguito per conflitto con tb_servizi dell'applicativo presenze.
'sql = Aggiornamento__FRAMEWORK_CORE__110(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 290)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 291
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__111(conn)
CALL DB.Execute(sql, 291)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 292
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__112(conn)
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
sql = Aggiornamento__FRAMEWORK_CORE__113(conn)
CALL DB.Execute(sql, 293)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 294
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__114(DB, rs, 294)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 295
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 295)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 296
'riaggiorna la vista per dell'aggiornamento 115
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 296)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 297
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__116(conn)
'	  Aggiornamento__FRAMEWORK_CORE__116_bis(conn):		non eseguito per conflitto con tb_servizi dell'applicativo presenze.
CALL DB.Execute(sql, 297)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 298
'...........................................................................................
'non eseguito per conflitto con tb_servizi dell'applicativo presenze.
'sql = Aggiornamento__FRAMEWORK_CORE__117(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 298)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 299
'...........................................................................................
'non eseguito per conflitto con tb_servizi dell'applicativo presenze.
'sql = Aggiornamento__FRAMEWORK_CORE__118(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 299)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 300
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__119(conn)
CALL DB.Execute(sql, 300)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 301
'...........................................................................................
'modifica vista disponibiilt&ograve; alberghi unindustria / turive
'...........................................................................................
sql = DropObject(conn, "vwp_disponibilita", "VIEW")
CALL DB.Execute(sql, 301)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(301)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 302 
'...........................................................................................
' Installa booking3 per gestire le strutture ricettive e la disponibilità
sql = Install__BOOKING3(conn)
CALL DB.Execute(sql, 302)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 303 
'...........................................................................................
' Esegue tutti gli aggiornamenti da 1 a 32 per allineare il booking alla versione corrente
dim i
for i = 1 to 32
	sql = Eval("Aggiornamento__BOOKING3__"&i&"(conn)")
	CALL DB.Execute(sql, 302+i)
	CALL DB.ReSyncTransaction()
next
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 335 
'...........................................................................................
' Ricominciare da qui ...
sql = Aggiornamento__BOOKING3__33(conn)
CALL DB.Execute(sql, 335)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 336
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__120(conn)
CALL DB.Execute(sql, 336)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 337
'...........................................................................................
'non eseguito per conflitto con tb_servizi dell'applicativo presenze.
sql = "ALTER TABLE vtb_strutture ADD  " & vbCrLf & _
	  "str_codiceregionale NVARCHAR(20); " & vbCrLf
CALL DB.Execute(sql, 337)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 338
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__121(DB, rs, 338)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 339
'...........................................................................................
' .
sql = "ALTER TABLE tb_servizi " & vbCrLf & _
	  "ADD codice_giust NCHAR(2);" & vbCrLf
CALL DB.Execute(sql, 339)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 340
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__122(conn)
CALL DB.Execute(sql, 340)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 341
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__123(conn)
CALL DB.Execute(sql, 341)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 342
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__124(conn)
CALL DB.Execute(sql, 342)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 343
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__125(conn)
CALL DB.Execute(sql, 343)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 344
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__126(conn)
CALL DB.Execute(sql, 344)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 345
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__127(conn)
CALL DB.Execute(sql, 345)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 346
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__128(conn)
CALL DB.Execute(sql, 346)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 347
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__129(conn)
CALL DB.Execute(sql, 347)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 348
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__130(conn)
CALL DB.Execute(sql, 348)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(348)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 349
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__131(conn)
CALL DB.Execute(sql, 349)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 350
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__132(conn)
CALL DB.Execute(sql, 350)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 351
'...........................................................................................
' .
sql = "ALTER TABLE tb_servizi " & vbCrLf & _
	  "ADD se_in_statistica bit;" & vbCrLf
CALL DB.Execute(sql, 351)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 352
'...........................................................................................
' .
sql = "CREATE PROCEDURE Get_Stat_Assenze" & vbCrLf & _
	  "    @id_admin int," & vbCrLf & _
	  "	   @anno_stat int" & vbCrLf & _
	  "AS" & vbCrLf & _
	  "BEGIN" & vbCrLf & _
	  "   SELECT     tb_servizi.nome_servizio, tb_tipiServizi.nomeTipoServ, SUM(1 + ABS(DATEDIFF(dd, tb_richieste.data_inizio, tb_richieste.data_fine))) AS gg," & vbCrLf & _
	  "				    SUM(tb_richieste.valore_richiesta) AS giorni_richiesti" & vbCrLf & _
	  "	   FROM         tb_admin INNER JOIN" & vbCrLf & _
	  "					   tb_richieste ON tb_admin.id_admin = tb_richieste.id_dipendente INNER JOIN" & vbCrLf & _
	  "					   tb_servizi ON tb_richieste.tipo_richiesta = tb_servizi.id_servizio INNER JOIN" & vbCrLf & _
	  "					   tb_tipiServizi ON tb_servizi.tipo_servizio = tb_tipiServizi.idTipoServ" & vbCrLf & _
	  "	   WHERE     (tb_servizi.se_in_statistica = 1) AND (tb_richieste.tipo_messaggio LIKE 'RI' OR" & vbCrLf & _
      "               tb_richieste.tipo_messaggio LIKE 'N') AND (YEAR(tb_richieste.data_richiesta) = @anno_stat) AND (tb_admin.id_admin = @id_admin)" & vbCrLf & _
	  "	   GROUP BY tb_servizi.nome_servizio, tb_tipiServizi.nomeTipoServ" & vbCrLf & _
	  "	   ORDER BY tb_tipiServizi.nomeTipoServ, tb_servizi.nome_servizio" & vbCrLf & _
	  "END" & vbCrLf
CALL DB.Execute(sql, 352)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 353
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__133(conn)
CALL DB.Execute(sql, 353)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(353)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 353
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__134(conn)
CALL DB.Execute(sql, 353)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 354
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__135(conn)
CALL DB.Execute(sql, 354)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 355
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__136(conn)
CALL DB.Execute(sql, 355)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 356
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__137(conn)
CALL DB.Execute(sql, 356)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(356)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 357
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__138(conn)
CALL DB.Execute(sql, 357)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 358
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__139(conn)
CALL DB.Execute(sql, 358)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 359
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__140(conn)
CALL DB.Execute(sql, 359)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 360
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__141(conn)
CALL DB.Execute(sql, 360)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(360)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 361
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__142(conn)
CALL DB.Execute(sql, 361)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(361)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 362
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__143(conn)
CALL DB.Execute(sql, 362)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(362)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 363
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__144(conn)
CALL DB.Execute(sql, 363)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(363)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 364
'...........................................................................................
sql = Install__Questionario(conn)
CALL DB.Execute(sql, 364)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 365
'...........................................................................................
sql = Aggiornamento__Questionario__1(conn)
CALL DB.Execute(sql, 365)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 366
'...........................................................................................
sql = Aggiornamento__Questionario__2(conn)
CALL DB.Execute(sql, 366)
'*******************************************************************************************



'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 367
'...........................................................................................
sql = " CREATE TABLE dbo.tb_iat ( " + _
	  "		iat_id " & SQL_PrimaryKeyInt(conn, "tb_iat") + ", " + _
	  "		iat_nome " + SQL_CharField(Conn, 250) + ", " + _
	  "		iat_saldo_cassa money NULL, " + _
	  "		iat_fondo_cassa money NULL); " + _
	  " CREATE TABLE dbo.rel_iat_admin ( " + _
	  "		rel_id " & SQL_PrimaryKeyInt(conn, "rel_iat_admin") + ", " + _
	  "		rel_iat_id INT NULL, " + _
	  "		rel_admin_id INT NULL); " + _
	  SQL_AddForeignKey(conn, "rel_iat_admin", "rel_iat_id", "tb_iat", "iat_id", true, "") + _
	  SQL_AddForeignKey(conn, "rel_iat_admin", "rel_admin_id", "tb_admin", "id_admin", true, "")

CALL DB.Execute(sql, 367)

if DB.last_update_executed then
	Dim MAG_conn, rs_DATA, rs_MAG
	
	set MAG_conn = Server.CreateObject("ADODB.Connection")
	MAG_conn.open Application("MAG_IAT_ConnectionString")
	
	set rs_DATA = Server.CreateObject("ADODB.RecordSet")
	set rs_MAG = Server.CreateObject("ADODB.RecordSet")

	sql = " SELECT * FROM tb_iat "
	rs_DATA.open sql, conn, adOpenDynamic, adLockOptimistic
	rs_MAG.open sql, MAG_conn, adOpenStatic, adLockOptimistic
	while not rs_MAG.eof 
		rs_DATA.AddNew
		rs_DATA("iat_id") = rs_MAG("iat_id")
		rs_DATA("iat_nome") = rs_MAG("iat_nome")
		rs_DATA("iat_saldo_cassa") = rs_MAG("iat_saldo_cassa")
		rs_DATA("iat_fondo_cassa") = rs_MAG("iat_fondo_cassa")
		rs_DATA.Update
		rs_MAG.moveNext
	wend
	rs_DATA.close
	rs_MAG.close
	
	sql = " SELECT * FROM rel_iat_admin "
	rs_DATA.open sql, conn, adOpenKeyset, adLockOptimistic
	rs_MAG.open sql, MAG_conn, adOpenStatic, adLockReadOnly, adAsyncFetch
	while not rs_MAG.eof 
		rs_DATA.AddNew
		rs_DATA("rel_id") = rs_MAG("rel_id")
		rs_DATA("rel_iat_id") = rs_MAG("rel_iat_id")
		rs_DATA("rel_admin_id") = rs_MAG("rel_admin_id")
		rs_DATA.Update
		rs_MAG.moveNext
	wend
	rs_DATA.close
	rs_MAG.close
	set rs_DATA = nothing
	set rs_MAG = nothing
end if
'*******************************************************************************************
'*******************************************************************************************


'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 368
'...........................................................................................
sql = " ALTER TABLE qtb_compilazioni " + _
	  "	ADD	comp_iat_id INT NULL; " + _
	  SQL_AddForeignKey(conn, "qtb_compilazioni", "comp_iat_id", "tb_iat", "iat_id", false, "") + _
	  " CREATE TABLE dbo.rel_iat_questionari ( " + _
	  "		riq_id " & SQL_PrimaryKey(conn, "rel_iat_questionari") + ", " + _
	  "		riq_iat_id INT NULL, " + _
	  "		riq_questionario_id INT NULL); " + _
	  SQL_AddForeignKey(conn, "rel_iat_questionari", "riq_iat_id", "tb_iat", "iat_id", true, "") + _
	  SQL_AddForeignKey(conn, "rel_iat_questionari", "riq_questionario_id", "qtb_questionari", "quest_id", true, "")

CALL DB.Execute(sql, 368)
'*******************************************************************************************
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 369
'...........................................................................................
sql = Aggiornamento__Questionario__3(conn)
CALL DB.Execute(sql, 369)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 370
'...........................................................................................
sql = Aggiornamento__Questionario__4(conn)
CALL DB.Execute(sql, 370)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 371
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__145(conn)
CALL DB.Execute(sql, 371)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 372
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "ru")
CALL DB.Execute(sql, 372)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "ru", "russo", "Русский")
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 373
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "cn")
CALL DB.Execute(sql, 373)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "cn", "Cinese", "中文")
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 374
'...........................................................................................
sql = Aggiornamento__Questionario__5(conn)
CALL DB.Execute(sql, 374)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 375
'...........................................................................................
sql = Aggiornamento__Questionario__6(conn)
CALL DB.Execute(sql, 375)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 376
'...........................................................................................
sql = Aggiornamento__Questionario__7(conn)
CALL DB.Execute(sql, 376)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(376)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 377
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "pt")
CALL DB.Execute(sql, 377)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "pt", "Portoghese", "Português")
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(377)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 378
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__148(conn)
CALL DB.Execute(sql, 378)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 379
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__149(conn)
CALL DB.Execute(sql, 379)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(379)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 380
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__150(conn)
CALL DB.Execute(sql, 380)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(380)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 381
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__151(conn)
CALL DB.Execute(sql, 381)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 382
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__152(conn)
CALL DB.Execute(sql, 382)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 383
'...........................................................................................
'aggiungi colonna per la gestione degli ambiti la tabella ambiti
'...........................................................................................
sql = " CREATE TABLE dbo.tb_ambiti_territoriali ( " + _
	  "		ambito_id " & SQL_PrimaryKey(conn, "tb_ambiti_territoriali") + ", " + _
	  "		ambito_nome " + SQL_CharField(Conn, 250) + ", " + _
	  "     ambito_note " + SQL_CharField(Conn, 0) +" NULL); " + _ 
	  " ALTER TABLE dbo.tb_iat " + _
	  "		ADD	iat_ambito_id INT NOT NULL DEFAULT 0; " + _
	  SQL_AddForeignKey(conn, "tb_iat", "iat_ambito_id", "tb_ambiti_territoriali", "ambito_id", false, "") + _
	  " INSERT INTO dbo.tb_ambiti_territoriali (ambito_nome,ambito_note) VALUES ('Venezia Mestre','');" + _
	  " INSERT INTO dbo.tb_ambiti_territoriali (ambito_nome,ambito_note) VALUES ('Jesolo Eraclea','');" + _
	  " INSERT INTO dbo.tb_ambiti_territoriali (ambito_nome,ambito_note) VALUES ('Bibione Caorle','');" + _
	  " INSERT INTO dbo.tb_ambiti_territoriali (ambito_nome,ambito_note) VALUES ('Chioggia Sottomarina','');" + _
	  " INSERT INTO dbo.tb_ambiti_territoriali (ambito_nome,ambito_note) VALUES ('Cavallino Treporti','');"
	  '	response.write sql
CALL DB.Execute(sql, 383)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 384
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__153(conn)
CALL DB.Execute(sql, 384)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(384)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 385
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__154(conn)
CALL DB.Execute(sql, 385)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 386
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__155(conn)
CALL DB.Execute(sql, 386)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 387
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__156(conn)
CALL DB.Execute(sql, 387)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(387)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 388
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__157(conn)
CALL DB.Execute(sql, 388)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(388)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 389
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__158(conn)
CALL DB.Execute(sql, 389)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 390
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__159(conn)
CALL DB.Execute(sql, 390)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(390)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 391
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__160(conn)
CALL DB.Execute(sql, 391)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(391)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 392
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__161(conn)
CALL DB.Execute(sql, 392)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 393
'...........................................................................................
sql = Aggiornamento__Questionario__8(conn)
CALL DB.Execute(sql, 393)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 394
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__162(conn)
CALL DB.Execute(sql, 394)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__162(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 395
'...........................................................................................
sql = Install__MEMO2(conn)
CALL DB.Execute(sql,395)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 396
'...........................................................................................
sql = AggiornamentoSpeciale__MEMO2__1(DB, rs, 396)
CALL DB.Execute(sql, 396)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 397
' Rimozione permessi da vecchi applicativi
'...........................................................................................
sql = 	"UPDATE rel_admin_sito " + _
		"SET sito_id = 26 WHERE sito_id = 2; " + _
		"DELETE " + _
		"FROM rel_admin_sito " + _
		"WHERE sito_id IN (102,104,107,108,110,111,126); " + _
		"DELETE FROM tb_siti WHERE id_sito IN (102,104,107,108,110,111,126);" + _
		""
CALL DB.Execute(sql, 397)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 398
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__163(conn)
CALL DB.Execute(sql, 398)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 399
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__165(conn)
CALL DB.Execute(sql, 399)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 400
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__166(conn)
CALL DB.Execute(sql, 400)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 401
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__167(conn)
CALL DB.Execute(sql, 401)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 402
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__168(conn)
CALL DB.Execute(sql, 402)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 403
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__169(conn)
CALL DB.Execute(sql, 403)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(403)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 404
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__170(conn)
CALL DB.Execute(sql, 404)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(404)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 405
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__171(conn)
CALL DB.Execute(sql, 405)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(405)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 406
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__172(conn)
CALL DB.Execute(sql, 406)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(406)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 407
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__173(conn)
CALL DB.Execute(sql, 407)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(407)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 408
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__174(conn)
CALL DB.Execute(sql, 408)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(408)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 409
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__175(conn)
CALL DB.Execute(sql, 409)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__175(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 410
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__176(conn)
CALL DB.Execute(sql, 410)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(410)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 411
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__177(conn)
CALL DB.Execute(sql, 411)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(411)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 412
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__178(conn)
CALL DB.Execute(sql, 412)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(412)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 413
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__179(conn)
CALL DB.Execute(sql, 413)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 414
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__180(conn)
CALL DB.Execute(sql, 414)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 414
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__181(conn)
CALL DB.Execute(sql, 414)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 415
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__182(conn)
CALL DB.Execute(sql, 415)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 416
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__183(conn)
CALL DB.Execute(sql, 416)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 417
'...........................................................................................
sql = Aggiornamento__MEMO2__1(conn)
CALL DB.Execute(sql, 417)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 418
'...........................................................................................
sql = Aggiornamento__MEMO2__2(conn)
CALL DB.Execute(sql, 418)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__2(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 419
'...........................................................................................
sql = Aggiornamento__MEMO2__3(conn)
CALL DB.Execute(sql, 419)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 420
'...........................................................................................
sql = Aggiornamento__MEMO2__4(conn)
CALL DB.Execute(sql, 420)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__4(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 421
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__184(conn)
CALL DB.Execute(sql, 421)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 422
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__185(conn)
CALL DB.Execute(sql, 422)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 423
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__186(conn)
CALL DB.Execute(sql, 423)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 424
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__187(conn)
CALL DB.Execute(sql, 424)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 425
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__188(conn)
CALL DB.Execute(sql, 425)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 426
'...........................................................................................
sql = Aggiornamento__MEMO2__5(conn)
CALL DB.Execute(sql, 426)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__5(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 427
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__189(conn)
CALL DB.Execute(sql, 427)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 428
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__190(conn)
CALL DB.Execute(sql, 428)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 429
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__191(conn)
CALL DB.Execute(sql, 429)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(429)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 430
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__192(conn)
CALL DB.Execute(sql, 430)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(430)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 431
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__193(conn)
CALL DB.Execute(sql, 431)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(431)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 432
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__194(conn)
CALL DB.Execute(sql, 432)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__194(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 433
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__195(conn)
CALL DB.Execute(sql, 433)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(433)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 434
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__196(conn)
CALL DB.Execute(sql, 434)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(434)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 435
'...........................................................................................
sql = Aggiornamento__MEMO2__6(conn)
CALL DB.Execute(sql, 435)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 436
'...........................................................................................
sql = Aggiornamento__MEMO2__7(conn)
CALL DB.Execute(sql, 436)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__7(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 437
'...........................................................................................
sql = Aggiornamento__MEMO2__8(conn)
CALL DB.Execute(sql, 437)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__MEMO2__8(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 438
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__197(conn)
CALL DB.Execute(sql, 438)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 439
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__198(conn)
CALL DB.Execute(sql, 439)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 440
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__199(conn)
CALL DB.Execute(sql, 440)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 441
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__200(conn)
CALL DB.Execute(sql, 441)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 442
'...........................................................................................
sql = Aggiornamento__MEMO2__9(conn)
CALL DB.Execute(sql, 442)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 443
'...........................................................................................
sql = Aggiornamento__MEMO2__10(conn)
CALL DB.Execute(sql, 443)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 444
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__201(conn)
CALL DB.Execute(sql, 444)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 445
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "ru")
CALL DB.Execute(sql, 445)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 446
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "pt")
CALL DB.Execute(sql, 446)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 447
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "cn")
CALL DB.Execute(sql, 447)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 448
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__203(conn)
CALL DB.Execute(sql, 448)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__203(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 449
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__204(conn)
CALL DB.Execute(sql, 449)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__204(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 450
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__205(conn)
CALL DB.Execute(sql, 450)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 451
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__206(conn)
CALL DB.Execute(sql, 451)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(451)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 452
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__207(conn)
CALL DB.Execute(sql, 452)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(452)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 453
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__208(conn)
CALL DB.Execute(sql, 453)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 454
'...........................................................................................
sql = Aggiornamento__MEMO2__11(conn)
CALL DB.Execute(sql, 454)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 455
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__209(conn)
CALL DB.Execute(sql, 455)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__209(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(455)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 456
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__210(conn)
CALL DB.Execute(sql, 456)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 457
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__211(conn)
CALL DB.Execute(sql, 457)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(457)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 458
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__212(conn)
CALL DB.Execute(sql, 458)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 459
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__213(conn)
CALL DB.Execute(sql, 459)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 460
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__214(conn)
CALL DB.ProtectedExecuteRebuild(sql, 460, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 461
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__215(conn)
CALL DB.ProtectedExecuteRebuild(sql, 461, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 462
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__216(conn)
CALL DB.ProtectedExecuteRebuild(sql, 462, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 463
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__217(conn)
CALL DB.Execute(sql, 463)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__217(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 464
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__218(conn)
CALL DB.ProtectedExecuteRebuild(sql, 464, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 465
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__219(conn)
CALL DB.ProtectedExecuteRebuild(sql, 465, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 466
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__220(conn)
CALL DB.ProtectedExecuteRebuild(sql, 466, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 467
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__221(conn)
CALL DB.ProtectedExecuteRebuild(sql, 467, false, true)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 468
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__222(conn)
CALL DB.Execute(sql, 468)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 469
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__223(conn)
CALL DB.ProtectedExecuteRebuild(sql, 469, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 470
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__224(conn)
CALL DB.Execute(sql, 470)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__224(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 471
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__225(conn)
CALL DB.ProtectedExecuteRebuild(sql, 471, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 472
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__226(conn)
CALL DB.ProtectedExecuteRebuild(sql, 472, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 473
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__227(conn)
CALL DB.ProtectedExecuteRebuild(sql, 473, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 474
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__228(conn)
CALL DB.ProtectedExecuteRebuild(sql, 474, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 475
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__229(conn)
CALL DB.ProtectedExecuteRebuild(sql, 475, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 476
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__230(conn)
CALL DB.Execute(sql, 476)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__230(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 477
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__231(conn)
CALL DB.ProtectedExecuteRebuild(sql, 477, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 478
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__232(conn)
CALL DB.ProtectedExecuteRebuild(sql, 478, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 479
'...........................................................................................
sql = Aggiornamento__MEMO2__12(conn)
CALL DB.Execute(sql, 479)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 480
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__233(conn)
CALL DB.ProtectedExecuteRebuild(sql, 480, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 481
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__234(conn)
CALL DB.ProtectedExecuteRebuild(sql, 481, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 482
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__235(conn)
CALL DB.ProtectedExecuteRebuild(sql, 482, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 483
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__236(conn)
CALL DB.ProtectedExecuteRebuild(sql, 483, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 484
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__237(conn)
CALL DB.ProtectedExecuteRebuild(sql, 484, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 485
'...........................................................................................
sql = Aggiornamento__MEMO2__13(conn)
CALL DB.Execute(sql, 485)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 486
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__238(conn)
CALL DB.ProtectedExecuteRebuild(sql, 486, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 487
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__239(conn)
CALL DB.ProtectedExecuteRebuild(sql, 487, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 488
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__240(conn)
CALL DB.ProtectedExecuteRebuild(sql, 488, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 489
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__241(conn)
CALL DB.ProtectedExecuteRebuild(sql, 489, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 490
'...........................................................................................
'aggiorna campi parametri applicativi
'...........................................................................................
sql = " ALTER TABLE qtb_domande ADD " + _
	  "		dom_obbligatoria bit NULL "
CALL DB.Execute(sql, 490)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 491
'...........................................................................................
sql = Aggiornamento__MEMO2__14(conn)
CALL DB.Execute(sql, 491)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 492
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__242(conn)
CALL DB.ProtectedExecuteRebuild(sql, 492, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 493
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__243(conn)
CALL DB.ProtectedExecuteRebuild(sql, 493, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 494
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__244(conn)
CALL DB.ProtectedExecuteRebuild(sql, 494, false, true)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
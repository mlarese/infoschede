<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>


<!--#INCLUDE VIRTUAL="admin/tools.asp" -->
<!--#INCLUDE VIRTUAL="admin/Tools4Save.asp" -->
<!--#INCLUDE VIRTUAL="admin/constant_definition.asp" -->
<!--#INCLUDE FILE="subscripts/Update_turismoprovincia_import_tools.asp" -->
<!--#INCLUDE FILE="subscripts/Update_turismoprovincia_updateimport_tools.asp" -->
<!--#INCLUDE FILE="../../NextGallery/Tools_Gallery.asp" -->
<%
'...........................................................................................
set index.conn = conn
'...........................................................................................


'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************
'OGGETTI e VARIABILI GLOBALI per aggiornamento dati: GESTIONE ALBERI AREE e CATEGORIE VARIE
set Index.Conn = conn
set iCatAnagrafiche.conn = conn
set iCatEventi.conn = conn
set iAree.conn = conn
set CatGallery.conn = conn
'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 223
'...........................................................................................
'rinomina la tabella tb_news per non interferire con quella del framework
'...........................................................................................
sql = " CREATE TABLE dbo.tb_turismo_news ( " + _
	  "		news_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		news_titolo nvarchar (255) NULL , " + _
	  "		news_sottotitolo nvarchar (255) NULL , " + _
	  "		news_testo ntext NULL , " + _
	  "		news_dtpubbl smalldatetime NULL , " + _
	  "		news_dtscadenza smalldatetime NOT NULL " + _
	  " ); " + _
	  " ALTER TABLE tb_turismo_news WITH NOCHECK ADD CONSTRAINT PK_tb_turismo_news PRIMARY KEY NONCLUSTERED (news_id); " + _
	  " INSERT INTO tb_turismo_news(news_titolo, news_sottotitolo, news_testo, news_dtpubbl, news_dtscadenza) " + _
	  " SELECT news_titolo, news_sottotitolo, news_testo, news_dtpubbl, news_dtscadenza FROM tb_news; " + _
	  " DROP TABLE tb_news; "
CALL DB.Execute(sql, 223)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 224
'...........................................................................................
'modifica del passport gia presente per uniformarlo al framework
'...........................................................................................
sql = " CREATE TABLE dbo.tb_turismo_admin ( " + _
	  "		id_admin int NOT NULL , " + _
	  "		admin_apt nvarchar (2) NULL " + _
	  " ); " + _
	  " INSERT INTO tb_turismo_admin(id_admin, admin_apt) " + _
	  " SELECT id_admin, admin_apt FROM tb_admin; " + _
	  " ALTER TABLE tb_admin DROP COLUMN admin_apt; " + _
	  " CREATE TRIGGER dbo.tr__tb_admin__insert ON tb_admin FOR INSERT AS " + _
	  " INSERT INTO tb_turismo_admin(id_admin) SELECT id_admin FROM inserted; " + _
	  " CREATE TRIGGER dboON tb_admin FOR DELETE AS " + _
	  " DELETE FROM tb_turismo_admin WHERE id_admin IN (SELECT id_admin FROM deleted); " + _
	  " ALTER TABLE tb_admin ADD " + _
	  "		admin_matricola nvarchar (50) NULL , " + _
	  "		admin_data_Nasc smalldatetime NULL , " + _
      "		admin_data_assunz smalldatetime NULL , " + _
	  "		admin_note text NULL , " + _
	  "		admin_ufficio int NULL , " + _
	  "		admin_contratto int NULL , " + _
	  "		admin_direttore bit NULL , " + _
	  "		admin_contatto int NULL , " + _
	  "		admin_profilo int NULL , " + _
	  "		admin_scadenza smalldatetime NULL; " + _
	  " ALTER TABLE log_admin DROP CONSTRAINT FK_log_admin_tb_admin; " + _
	  " ALTER TABLE log_admin DROP CONSTRAINT FK_log_admin_tb_siti; " + _
	  " ALTER TABLE log_admin ADD " + _
	  "		log_admin_id int NULL , " + _
	  "		log_sito_id int NULL; " + _
	  " UPDATE log_admin SET log_admin_id = log_id_admin, log_sito_id = log_id_sito; " + _
	  " ALTER TABLE log_admin DROP COLUMN log_id_sito; " + _
	  " ALTER TABLE log_admin DROP COLUMN log_id_admin; " + _
	  " ALTER TABLE log_admin ADD " + _
	  "		CONSTRAINT FK_log_admin_tb_admin FOREIGN KEY (log_admin_id) REFERENCES tb_admin (id_admin) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE, " + _
	  "		CONSTRAINT FK_log_admin_tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE tb_siti DROP COLUMN sito_corrente; " + _
	  " ALTER TABLE tb_siti ADD " + _
	  "		sito_amministrazione bit NULL , " + _
	  "		sito_rubrica_area_riservata int NULL; " + _
	  " UPDATE tb_siti SET sito_amministrazione = 1; " + _
	  " CREATE TABLE dbo.tb_siti_parametri ( " + _
	  "		par_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		par_key nvarchar (250) NULL , " + _
	  "		par_value nvarchar (250) NULL , " + _
	  "		par_sito_id int NOT NULL " + _
	  " ); " + _
	  " ALTER TABLE tb_siti_parametri WITH NOCHECK ADD CONSTRAINT PK_tb_siti_parametri PRIMARY KEY NONCLUSTERED (par_id); " + _
	  " CREATE TABLE dbo.tb_Utenti ( " + _
	  "		ut_ID int IDENTITY (1, 1) NOT NULL , " + _
	  "		ut_NextCom_ID int NOT NULL , " + _
	  " 	ut_login nvarchar (50) NULL , " + _
	  "		ut_password nvarchar (50) NULL , " + _
	  "		ut_Abilitato bit NULL , " + _
	  "		ut_ScadenzaAccesso smalldatetime NULL " + _
	  " ); " + _
	  " ALTER TABLE tb_Utenti WITH NOCHECK ADD CONSTRAINT PK_tb_Utenti PRIMARY KEY NONCLUSTERED (ut_ID); " + _
	  " CREATE TABLE dbo.rel_utenti_sito ( " + _
	  "		rel_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		rel_ut_id int NOT NULL , " + _
	  "		rel_sito_id int NOT NULL , " + _
	  "		rel_permesso int NOT NULL " + _
	  " ); " + _
	  " ALTER TABLE rel_utenti_sito WITH NOCHECK ADD CONSTRAINT PK_rel_utenti_sito PRIMARY KEY NONCLUSTERED (rel_id); " + _
	  " CREATE TABLE dbo.log_utenti ( " + _
	  "		log_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		log_ut_id int NOT NULL , " + _
	  "		log_sito_id int NOT NULL , " + _
	  "		log_data smalldatetime NULL , " + _
	  "		log_username nvarchar (50) NULL " + _
	  " ); " + _
	  " ALTER TABLE log_utenti WITH NOCHECK ADD CONSTRAINT PK_log_utenti PRIMARY KEY NONCLUSTERED (log_id); " + _
	  " ALTER TABLE log_utenti ADD " + _
	  "		CONSTRAINT FK_log_utenti__tb_utenti FOREIGN KEY (log_ut_id) REFERENCES tb_Utenti (ut_ID) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE, " + _
	  "		CONSTRAINT FK_log_utenti__tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE rel_utenti_sito ADD " + _
	  "		CONSTRAINT FK_rel_utenti_sito__tb_utenti FOREIGN KEY (rel_ut_id) REFERENCES tb_Utenti (ut_ID) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE, " + _
	  "		CONSTRAINT FK_rel_utenti_sito__tb_siti FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE tb_siti_parametri ADD " + _
	  "		CONSTRAINT FK_tb_siti_parametri_tb_siti FOREIGN KEY (par_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; "
CALL DB.Execute(sql, 224)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 225
'...........................................................................................
'installazione applicativi del framework core rimanenti
'...........................................................................................
sql = Install__FRAMEWORK_CORE__NEXTCOM(conn) + _
	  Install__FRAMEWORK_CORE__NEXTLINK(conn) + _
	  Install__FRAMEWORK_CORE__NEXTNEWS(conn) + _
	  Install__FRAMEWORK_CORE__NEXTGALLERY(conn) + _
	  Install__FRAMEWORK_CORE__NEXTFAQ(conn) + _
	  Install__FRAMEWORK_CORE__NEXTTEAM(conn)
CALL DB.Execute(sql, 225)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 226
'...........................................................................................
'tolgo l'identity alla chiave di tb_siti
'...........................................................................................
sql = " ALTER TABLE rel_utenti_sito DROP CONSTRAINT FK_rel_utenti_sito__tb_siti; " + _
	  " ALTER TABLE rel_admin_sito DROP CONSTRAINT FK_rel_admin_sito_tb_siti; " + _
	  " ALTER TABLE log_admin DROP CONSTRAINT FK_log_admin_tb_siti; " + _
	  " ALTER TABLE tb_siti_parametri DROP CONSTRAINT FK_tb_siti_parametri_tb_siti; " + _
	  " ALTER TABLE log_utenti DROP CONSTRAINT FK_log_utenti__tb_siti; " + _
	  " ALTER TABLE tb_siti DROP CONSTRAINT PK_tb_siti; " + _
	  " ALTER TABLE tb_siti ADD " + _
	  "		id_sito2 int NULL; " + _
	  " UPDATE tb_siti SET id_sito2 = id_sito; " + _
	  " ALTER TABLE tb_siti DROP COLUMN id_sito; " + _
	  " ALTER TABLE tb_siti ADD " + _
	  "		id_sito int NOT NULL DEFAULT 0; " + _
	  " UPDATE tb_siti SET id_sito = id_sito2; " + _
	  " ALTER TABLE tb_siti DROP COLUMN id_sito2; " + _
	  " ALTER TABLE tb_siti WITH NOCHECK ADD CONSTRAINT PK_tb_siti PRIMARY KEY NONCLUSTERED (id_sito); " + _
	  " ALTER TABLE log_admin ADD " + _
	  "		CONSTRAINT FK_log_admin_tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE log_utenti ADD " + _
	  "		CONSTRAINT FK_log_utenti__tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE rel_admin_sito ADD " + _
	  "		CONSTRAINT FK_rel_admin_sito__tb_siti FOREIGN KEY (sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE rel_utenti_sito ADD " + _
	  "		CONSTRAINT FK_rel_utenti_sito__tb_siti FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; " + _
	  " ALTER TABLE tb_siti_parametri ADD " + _
	  "		CONSTRAINT FK_tb_siti_parametri_tb_siti FOREIGN KEY (par_sito_id) REFERENCES tb_siti (id_sito) " + _
	  "		ON DELETE CASCADE ON UPDATE CASCADE; "
CALL DB.Execute(sql, 226)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 227
'...........................................................................................
'corregge i parametri delle applicazioni esistenti
'...........................................................................................
sql = " UPDATE tb_siti SET id_sito = id_sito + 200; " + _
	  " UPDATE tb_siti SET sito_nome = 'NEXT-passport [gestione utenti]', sito_dir = 'NEXTpassport', sito_p1 = 'PASS_ADMIN', " + _
	  "					   sito_p2 = 'PASS_AMMINISTRATORI', sito_p3 = 'PASS_UTENTI', sito_rubrica_area_riservata = 0, id_sito = 1 " + _
	  " WHERE id_sito = 201; " + _
	  " UPDATE tb_siti SET sito_nome = 'NEXT-web 4.0 [gestione grafica e contenuti]', sito_dir = 'NEXTweb4', sito_p1 = 'WEB_ADMIN', " + _
	  "					   sito_p2 = 'WEB_USER', id_sito = 25 " + _
	  " WHERE id_sito = 202; "
CALL DB.Execute(sql, 227)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 228
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__1(conn)
CALL DB.Execute(sql, 228)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 229
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__2(conn)
CALL DB.Execute(sql, 229)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 230
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__3(conn)
CALL DB.Execute(sql, 230)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 231
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__4(conn)
CALL DB.Execute(sql, 231)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 232
'...........................................................................................
sql = rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
CALL DB.Execute(sql, 232)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 233
'...........................................................................................
CALL rebuild__FRAMEWORK_CORE__cartelle(conn, rs, DB, 233)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 234
'...........................................................................................
'rimuove tabelle di descrizione del database non piu' utilizzate
'...........................................................................................
sql = " ALTER TABLE _tb_properties DROP CONSTRAINT FK__tb_properties__tb_objects ;" + _
	  " drop table _tb_objects; " + _
	  " drop table _tb_properties; "
CALL DB.Execute(sql, 234)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 235
'...........................................................................................
'modifica e sostituisce la gestione delle tipologie per le agenzie di viaggio
' MOD_ID = 38 (agenzie di viaggio)
' tipi:		46 = sede principale
'			47 = filiale
'			48 = succursale
'...........................................................................................
sql = " UPDATE tb_strutture SET den_agg='sede' WHERE tipo=46; " + _
	  " UPDATE tb_strutture SET den_agg='filiale' WHERE tipo=47; " + _
	  " UPDATE tb_strutture SET den_agg='succursale' WHERE tipo=48; " + _
	  " INSERT INTO tb_tipi_str (tip_mod_id, tip_den_it, tip_valid_from, tip_cod_regione, tip_cod_rvt) " + _
	  " 	SELECT tip_mod_id, 'Agenzie di viaggio', tip_valid_from, tip_cod_regione, tip_cod_rvt " + _
	  " 	FROM tb_tipi_str WHERE tip_id=46; " + _
	  " UPDATE tb_strutture SET tipo=(SELECT MAX(tip_id) FROM tb_tipi_str WHERE tip_mod_id=38) WHERE (tipo=46 OR tipo=47 OR tipo=48); " + _
	  " INSERT INTO tb_tipi_str (tip_mod_id, tip_den_it, tip_valid_from, tip_cod_regione, tip_cod_rvt) " + _
	  " 	SELECT tip_mod_id, 'Associazioni ONLUS', tip_valid_from, tip_cod_regione, tip_cod_rvt " + _
	  " 	FROM tb_tipi_str WHERE tip_id=46; " + _
	  " DELETE FROM tb_tipi_str WHERE (tip_id=46 OR tip_id=47 OR tip_id=48) "
CALL DB.Execute(sql, 235)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 236
'...........................................................................................
'installa next-com e relativo permesso per amministratore
'...........................................................................................
sql = " INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3, sito_rubrica_area_riservata) " + _
	  " VALUES (3, 'NEXT-com [gestione comunicazioni]', 'NEXTcom', 1, 'COM_ADMIN', 'COM_USER', 'COM_POWER', 0); " + _
	  " INSERT INTO rel_admin_sito(admin_id, rel_as_permesso, sito_id) VALUES (1, 1, 3); "
CALL DB.Execute(sql, 236)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 237
'...........................................................................................
'corregge nomi applicazioni
'...........................................................................................
'sql = "UPDATE tb_siti SET sito_nome='Assessorato al turismo [gestione contenuti portale]' WHERE id_sito=" & TURISMO_CONTENT & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Assessorato al turismo [gesione moduli on-line]' WHERE id_sito=" & TURISMO_MODULI_OL & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Alberghi]' WHERE id_sito=" & TURISMO_ALBERGHI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Campeggi]' WHERE id_sito=" & TURISMO_CAMPEGGI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Affittacamere]' WHERE id_sito=" & TURISMO_AFFITTACAMERE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Unit&agrave;; abitative classificate]' WHERE id_sito=" & TURISMO_UA_CLASSIFICATE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Guide turistiche]' WHERE id_sito=" & TURISMO_GUIDE_TURISTICHE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Accompagnatori turistici]' WHERE id_sito=" & TURISMO_ACCOMPAGNATORI_TURISTICI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Guide naturalistico-ambientali]' WHERE id_sito=" & TURISMO_GUIDE_NATURALISTICHE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Animatori turistici]' WHERE id_sito=" & TURISMO_ANIMATORI_TURISTICI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Residence]' WHERE id_sito=" & TURISMO_RESIDENCE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Ricettivit&agrave;; sociali]' WHERE id_sito=" & TURISMO_RICETTIVITA_SOCIALI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Bed & breakfast]' WHERE id_sito=" & TURISMO_B_B & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Unit&agrave;; abitative non classificate]' WHERE id_sito=" & TURISMO_UA_NCLASSIFICATE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Foresterie]' WHERE id_sito=" & TURISMO_FORESTERIE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Country house]' WHERE id_sito=" & TURISMO_COUNTRY_HOUSE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Strutture ricettive [Unit&agrave;; abitative agenzie immobiliari]' WHERE id_sito=" & TURISMO_UA_AGENZIE & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Assessorato al turismo [ricerca comune]' WHERE id_sito=" & TURISMO_COMMON_SEARCH & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Direttori tecnici]' WHERE id_sito=" & TURISMO_DIRETTORI_TECNICI & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Agenzie di viaggio]' WHERE id_sito=" & TURISMO_AGENZIE_VIAGGIO & ";" + _
'	  "UPDATE tb_siti SET sito_nome='Professioni turistiche [Accompagnatori turistici agenzie]' WHERE id_sito=" & TURISMO_ACCOMPAGNATORI_AGENZIE & ";"
'CALL DB.Execute(sql, 237)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 238
'...........................................................................................
'aggiunge trigger per gestione rubriche collegate ai modelli
'aggiunge anche rubriche per la sincronizzazione
'...........................................................................................
sql = " CREATE TRIGGER dbo.tb_modelli_INSERT ON tb_modelli AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "				SELECT CASE mod_tipo_record " + vbCrLF + _
	  "					   	WHEN 'O' THEN 'Proprietari - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   	WHEN 'A' THEN 'Agenzie - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   	WHEN 'P' THEN 'Prof. Tur. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   	WHEN 'U' THEN 'U. A. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   	ELSE 'Strutture - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   END ,  " + vbCrLF + _
	  "					   'view_strutture', 'tb_modelli', mod_id, 1, 1 FROM INSERTED WHERE NOT(mod_tipo_record LIKE 'T')" & vbCrLf &_
	  "			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "				SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " & vbCrLf &_
	  " CREATE TRIGGER dbo.tb_modelli_UPDATE ON tb_modelli AFTER UPDATE AS " + vbCrLf + _
	  "		DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLf + _
	  "		DECLARE @MOD_ID int " + vbCrLF + _
	  "		SELECT TOP 1 @TIPO_RECORD = mod_tipo_record, @MOD_ID = mod_id FROM INSERTED " + vbCrLf + _
	  " 	IF @TIPO_RECORD = 'T' " + vbCrlf + _
	  "			BEGIN " + vbCrLf + _
	  "				IF UPDATE(mod_tipo_record) " + vbCrLF + _
	  "					BEGIN " + vbCrLf + _
	  "						UPDATE tb_rubriche SET locked_rubrica=0, rubrica_esterna=0 " + vbCrLf + _
	  "						WHERE SyncroFilterTable LIKE 'tb_modelli' AND SyncroFilterKey=@MOD_ID " + vbCrLf + _
	  "					END " + vbCrLf + _
	  "			END " + vbCrLf + _
	  "		ELSE " + vbCrLf + _
	  "			BEGIN " + vbCrLf + _
	  "				IF UPDATE(mod_strutture) " + vbCrLF + _
	  "					BEGIN " + vbCrLf + _
	  "						UPDATE tb_rubriche SET nome_rubrica = (SELECT TOP 1 CASE mod_tipo_record " + vbCrLF + _
	  "					   															WHEN 'O' THEN 'Proprietari - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   															WHEN 'A' THEN 'Agenzie - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   															WHEN 'P' THEN 'Prof. Tur. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   															WHEN 'U' THEN 'U. A. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   															ELSE 'Strutture - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "					   															END " + vbCrLF + _
	  "																				FROM INSERTED) " + vbCrLf + _
	  "						WHERE SyncroFilterTable='tb_modelli' AND SyncroFilterKey  = @MOD_ID " + vbCrLf + _
	  "					END " + vbCrLf + _
	  "			END; " + vbCrLf + _
	  " CREATE TRIGGER dbo.tb_modelli_DELETE ON tb_modelli AFTER DELETE AS " + vbCrLf + _
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable='tb_modelli' AND SyncroFilterKey IN (SELECT mod_id FROM DELETED) ; " + vbCrLf + _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "		SELECT CASE mod_tipo_record " + vbCrLF + _
	  "				WHEN 'O' THEN 'Proprietari - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				WHEN 'A' THEN 'Agenzie - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				WHEN 'P' THEN 'Prof. Tur. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				WHEN 'U' THEN 'U. A. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				ELSE 'Strutture - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   END , 'view_strutture', 'tb_modelli', mod_id, 1, 1 FROM tb_modelli WHERE NOT(mod_tipo_record LIKE 'T'); " + _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  " VALUES ('Globale (tutte le strutture e i professionisti)', 'view_strutture', 'tb_modelli', 0, 1, 1); " + vbCrLF + _
	  " INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) " + vbCrLf + _
	  "		SELECT id_rubrica, id_gruppo FROM tb_gruppi, tb_rubriche WHERE SyncroFilterTable LIKE 'tb_modelli' "
CALL DB.Execute(sql, 238)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 239
'...........................................................................................
'importa tutti i record delle strutture nel NEXT-com con le impostazioni di sincronizzazione
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 239)
if DB.last_update_executed then
	CALL Aggiornamento_239_SincronizzazioneStruttureNextCom(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_239_SincronizzazioneStruttureNextCom(DbConn, rs)
	dim readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.Open Application(request("ConnString")), "", ""
	
	sql = "SELECT RegCode, Modello, mod_tipo_record FROM VIEW_testata_strutture"
	readRs.open sql, readConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	while not readRs.eof %>
		<!-- <%= readRs("RegCode") %> -->
		<%CALL SincronizzaStruttura_NextCom(DbConn, rs, readRs("modello"), readRs("RegCode"), readRs("mod_tipo_record"))
		readRs.movenext
	wend	
	readRs.close
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 240
'...........................................................................................
'inserisce le rubriche per il collegamento tra associazioni e NEXT-com
'...........................................................................................
sql = "INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "		SELECT 'Assoc. - ' + nome_tipo , 'tb_assoc', 'tb_tipiassoc', id_tipo, 1, 1 FROM tb_tipiassoc; " + _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  " VALUES ('Globale (tutte le associazioni di categoria)', 'tb_assoc', 'tb_tipiassoc', 0, 1, 1); " + vbCrLF + _
	  " INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_Assegnato) " + vbCrLf + _
	  "		SELECT id_rubrica, id_gruppo FROM tb_gruppi, tb_rubriche WHERE SyncroFilterTable LIKE 'tb_assoc' "
CALL DB.Execute(sql, 240)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 241
'...........................................................................................
'importa tutti i record delle associazioni nel NEXT-com con le impostazioni di sincronizzazione
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 241)
if DB.last_update_executed then
	CALL Aggiornamento_241_SincronizzazioneAssociazioniNextCom(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_241_SincronizzazioneAssociazioniNextCom(DbConn, rs)
	dim readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.Open Application(request("ConnString")), "", ""
	
	sql = "SELECT asc_id FROM tb_assoc ORDER BY asc_id"
	readRs.open sql, readConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	while not readRs.eof
		CALL SincronizzaAssociazione_NextCom(DbConn, rs, readRs("asc_id"))
		readRs.movenext
	wend	
	readRs.close
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'...........................................................................................


'*******************************************************************************************
'AGGIORNAMENTO 242
'...........................................................................................
'cancello la stored procedure che elimina le vecchie news
'...........................................................................................
sql = "DROP PROCEDURE delete_news;"
CALL DB.Execute(sql, 242)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 243
'...........................................................................................
'aggiunge i trigger di gestione delle rubriche collegate alle categorie di associazioni
'...........................................................................................
sql = " CREATE TRIGGER dbo.tb_tipiassoc_INSERT ON tb_tipiassoc AFTER INSERT AS " & vbCrLf &_
	  "		INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " & vbCrLf &_
	  "				SELECT 'Assoc. - ' + nome_tipo , 'tb_assoc', 'tb_tipiassoc', id_tipo, 1, 1 FROM INSERTED" & vbCrLf &_
	  "			INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " & vbCrLf &_
	  "				SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " & vbCrLf &_
	  " CREATE TRIGGER dbo.tb_tipiassoc_UPDATE ON tb_tipiassoc AFTER UPDATE AS " + vbCrLf + _
	  "		IF UPDATE(nome_tipo) " + vbCrLF + _
	  "			BEGIN " + vbCrLf + _
	  "				UPDATE tb_rubriche SET nome_rubrica = (SELECT TOP 1 'Assoc. - ' + nome_tipo FROM INSERTED) " + vbCrLf + _
	  "				WHERE SyncroFilterTable='tb_tipiassoc' AND SyncroFilterKey = (SELECT TOP 1 id_tipo FROM INSERTED)" + vbCrLf + _
	  "			END; " + vbCrLf + _
	  " CREATE TRIGGER dbo.tb_tipiassoc_DELETE ON tb_tipiassoc AFTER DELETE AS " + vbCrLf + _
	  "		DELETE FROM tb_rubriche WHERE SyncroFilterTable='tb_tipiassoc' AND SyncroFilterKey IN (SELECT id_tipo FROM DELETED) ; "
CALL DB.Execute(sql, 243)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 244
'...........................................................................................
'aggiunge l'applicativo delle statistiche al passport
'...........................................................................................
sql = " INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_p1, sito_p2, sito_p3, sito_amministrazione) " + _
	  " 	VALUES (228, 'Assessorato al turismo [statistiche ed export dati]', '../statistiche', 'STAT_ADMIN', 'STAT_POWER', 'STAT_USER', 1); " + _
	  " INSERT INTO rel_admin_sito(admin_id, sito_id, rel_as_permesso) " + _
	  " 	VALUES (1, 228, 1) "
CALL DB.Execute(sql, 244)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 245
'...........................................................................................
'aggiunge colonne per la validazione dei record
'...........................................................................................
sql = " ALTER TABLE tb_loginStru ADD " + _
	  "		current_record_validato BIT NULL, " + _
	  "		avviso_inviato SMALLDATETIME NULL; " + _
	  " UPDATE tb_loginstru SET current_record_validato=0; " + _
	  " ALTER TABLE tb_loginStru ALTER COLUMN current_record_validato BIT NOT NULL; " + _
	  " ALTER TABLE tb_strutture ADD record_validato BIT NULL; " + _
	  " UPDATE tb_strutture SET record_validato=0; " + _
	  " ALTER TABLE tb_strutture ALTER COLUMN record_validato BIT NOT NULL ;"
CALL DB.Execute(sql, 245)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 246
'...........................................................................................
'modifica struttura dati per validazione record
'...........................................................................................
sql = " ALTER TABLE tb_loginStru DROP COLUMN current_record_validato; " + _
	  " ALTER TABLE tb_loginStru ADD " + _
	  "		current_valid_str_id INT NULL, " + _
	  "		DataValidazione SMALLDATETIME NULL ; " + _
	  " UPDATE tb_loginstru SET current_valid_str_id=current_str_id, DataValidazione=UltAgg; "
CALL DB.Execute(sql, 246)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 247
'...........................................................................................
'modifica struttura dati per validazione record
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD recod_validato_data SMALLDATETIME NULL; " + _
	  " UPDATE tb_strutture SET recod_validato_data = DataModifica, record_validato=1; "
CALL DB.Execute(sql, 247)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 248
'...........................................................................................
'modifica struttura dati per validazione record
'...........................................................................................
sql = " ALTER TABLE tb_strutture DROP COLUMN recod_validato_data; " + _
	  " ALTER TABLE tb_strutture ADD " + _
	  "		record_validato_data SMALLDATETIME NULL, " + _
	  "		record_validato_utente nvarchar(50) NULL, " + _
	  "		UtenteModifica nvarchar(50) NULL; " + _
	  " UPDATE tb_strutture SET " + _
	  "		UtenteModifica='SISTEMA', " + _
	  "		record_validato=1, " + _
	  "		record_validato_data = DataModifica, " + _
	  "		record_validato_utente = 'SISTEMA'; "
CALL DB.Execute(sql, 248)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 249
'...........................................................................................
'	modifica struttura dati per validazione record, avvisi di notifica password e validazione,
'	gestione periodi di dichiarazione
'...........................................................................................
sql = " ALTER TABLE tb_loginStru DROP COLUMN " + _
	  "				avviso_inviato, " + _
	  "				stampa_tb_prz_from, " + _
	  "				stampa_tb_prz_to " + _
	  "	; " + _
	  " ALTER TABLE tb_loginStru ADD " + _
	  "		avviso_inviato_data SMALLDATETIME NULL, " + _
	  "		avviso_inviato_utente nvarchar(50) " + _
	  " ; " + _
	  " ALTER TABLE tb_modelli ADD " + _
	  "		mod_esegue_dichiarazione_online BIT NULL, " + _
	  "		mod_page_email_conferma INT NULL, " + _
	  "		mod_page_email_validazione INT NULL, " + _
	  "		mod_page_email_password INT NULL " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_esegue_dichiarazione_online=1 " + _
	  "		WHERE mod_id IN (18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30, 32, 33, 34) ; " + _
	  " UPDATE tb_modelli SET mod_esegue_dichiarazione_online=0 WHERE mod_esegue_dichiarazione_online IS NULL; " + _ 
	  " UPDATE tb_modelli SET mod_page_email_password = mod_VMWEB_email_page; " + _
	  " ALTER TABLE tb_modelli ALTER COLUMN " + _
	  "		mod_esegue_dichiarazione_online BIT NOT NULL; " + _
	  " ALTER TABLE tb_modelli DROP COLUMN " + _
	  "		mod_print_inizio, " + _
	  " 	mod_print_fine, " + _
	  "		mod_VMWEB_email_page "
CALL DB.Execute(sql, 249)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 250
'...........................................................................................
'	rimuove colonne non necessarie da tb_loginstru e tb_strutture
'...........................................................................................
sql = " ALTER TABLE tb_loginstru DROP COLUMN " + _
	  "			DataValidazione, " + _
	  "			avviso_inviato_data, " + _
	  "			avviso_inviato_utente, " + _
	  "			ultagg, " + _
	  "			cancellata " + _
	  " ; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN " + _
	  "			DataModOL " + _
	  "	; "		
CALL DB.Execute(sql, 250)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 251
'...........................................................................................
'	aggiunge colonne per gestione validazione e dichiarazione online
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD" + _
	  " 	 avviso_inviato BIT NULL, " + _
	  "		 avviso_inviato_data SMALLDATETIME NULL, " + _
	  "		 avviso_inviato_utente nvarchar(50), " + _
	  "		 online_modifica_data SMALLDATETIME NULL, " + _
	  "		 online_modifica_utente nvarchar(50), " + _
	  "		 online_confermato BIT NULL, " + _
	  "		 online_confermato_data SMALLDATETIME NULL, " + _
	  "		 online_confermato_utente nvarchar(50), "  + _
	  "		 online_confermato_open_from SMALLDATETIME NULL, " + _
	  "		 online_confermato_open_to SMALLDATETIME NULL, " + _
	  "		 online_confermato_anno INT NULL, " + _
	  "		 online_confermato_periodo INT NULL "
CALL DB.Execute(sql, 251)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 252
'...........................................................................................
'	aggiorna tabella tb_loginStru
'...........................................................................................
sql = " ALTER TABLE tb_loginstru ADD " + _
	  "		current_DataModifica SMALLDATETIME NULL, " + _
	  "		current_valid_DataModifica SMALLDATETIME NULL "
CALL DB.Execute(sql, 252)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 253
'...........................................................................................
'	aggiunge a tabella modelli indicazione se modello esegue dichiarazione unica o meno
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "			mod_dichiarazione_unica BIT NULL, " + _
	  "			mod_dichiarazione_online BIT NULL ; " + _
	  " UPDATE tb_modelli SET mod_dichiarazione_online=1 WHERE mod_id IN (18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30, 32, 33); " + _
	  " UPDATE tb_modelli SET mod_dichiarazione_online=0 WHERE mod_dichiarazione_online IS NULL ; " + _
	  " UPDATE tb_modelli SET mod_dichiarazione_unica=1 WHERE mod_id IN (23, 24, 25, 26, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39) AND mod_dichiarazione_online=1; " + _
	  " UPDATE tb_modelli SET mod_dichiarazione_unica=0 WHERE mod_dichiarazione_unica IS NULL AND mod_dichiarazione_online=1; " + _
	  " ALTER TABLE tb_modelli ALTER COLUMN mod_dichiarazione_online BIT NOT NULL; " + _
	  " ALTER TABLE tb_modelli DROP CONSTRAINT DF__tb_modell__mod_a__603D47BB; " + _
	  " ALTER TABLE tb_modelli DROP CONSTRAINT DF__tb_modell__mod_p__61316BF4; " + _
	  " ALTER TABLE tb_modelli DROP COLUMN " + _
	  "			mod_esegue_dichiarazione_online, " + _
	  "			mod_anno_com, " + _
	  "			mod_periodo_com, " + _
	  "			mod_open_inizio, " + _
	  "			mod_open_fine "
CALL DB.Execute(sql, 253)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 254
'...........................................................................................
'	crea tabella per gestione dichiarazioni
'...........................................................................................
sql = " CREATE TABLE dbo.tb_dichiarazioni (" + _ 
	  "		dic_id int IDENTITY (1, 1) NOT NULL , " + _
	  "		dic_anno_prezzi INT NOT NULL, " + _
	  "		dic_data_inizio SMALLDATETIME NOT NULL, " + _
	  "		dic_data_fine SMALLDATETIME NOT NULL, " + _
	  " 	CONSTRAINT PK_tb_dichiarazione PRIMARY KEY (dic_id) " + _
	  ") ; " + _
	  " CREATE TABLE dbo.rel_dichiarazioni_modelli (" + _
	  "		rel_id int IDENTITY (1, 1) NOT NULL, " + _
	  "		rel_dic_id int NOT NULL, " + _
	  "		rel_mod_id INT NOT NULL, " + _
	  "		rel_tipo_dichiarazione INT NOT NULL, " + _
	  "		CONSTRAINT PK_rel_dichiarazioni_modelli PRIMARY KEY (rel_id), " + _
	  "		CONSTRAINT FK_rel_dichiarazioni_modelli_tb_dichiarazioni " + _
	  "			FOREIGN KEY (rel_dic_id) REFERENCES tb_dichiarazioni(dic_id) " + _
	  "			ON DELETE CASCADE  ON UPDATE CASCADE, " + _
	   "	CONSTRAINT FK_rel_dichiarazioni_modelli_tb_modelli " + _
	  "			FOREIGN KEY (rel_mod_id) REFERENCES tb_modelli(Mod_ID) " + _
	  "			ON DELETE CASCADE  ON UPDATE CASCADE " + _
	  ") ; "
CALL DB.Execute(sql, 254)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 255
'...........................................................................................
'	modifica tabella modelli per indicazione tipo dichiarazioni
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "		mod_dichiarazione_tipo nvarchar(1) NULL ; " + _
	  " ALTER TABLE tb_modelli DROP COLUMN mod_dichiarazione_unica; " + _
	  "	UPDATE tb_modelli SET mod_dichiarazione_tipo='" & DICHIARAZIONE_UNICA & "'   WHERE mod_id IN (			          23, 24, 25, 26,        29, 30,     32, 33, 34                    ); " + _
	  "	UPDATE tb_modelli SET mod_dichiarazione_tipo='" & DICHIARAZIONE_DOPPIA & "'  WHERE mod_id IN (18, 19, 20,     22,                 27, 28                                           ); " + _
	  "	UPDATE tb_modelli SET mod_dichiarazione_tipo='" & DICHIARAZIONE_NESSUNA & "' WHERE mod_id IN (            21,                                    31,            35, 36, 37, 38, 39 ); " + _
	  " ALTER TABLE tb_modelli ALTER COLUMN " + _
	  "		mod_dichiarazione_tipo nvarchar(1) NOT NULL ; "
CALL DB.Execute(sql, 255)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 256
'...........................................................................................
'	modifica disponibilita' tabella prezzi per modelli
'...........................................................................................
sql = " UPDATE tb_modelli SET mod_tabella_prezzi=0 WHERE mod_id IN (            21,     23, 24, 25, 26,         29, 30, 31, 32,     34, 35, 36, 37, 38, 39) " + _
	  " UPDATE tb_modelli SET mod_tabella_prezzi=1 WHERE mod_id IN (18, 19, 20,     22,                 27, 28,                 33                        ) "
CALL DB.Execute(sql, 256)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 257
'...........................................................................................
'	aggiunge campo per collegamento dichiarazione
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		online_dichiarazione_id INT NULL "
CALL DB.Execute(sql, 257)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 258
'...........................................................................................
'	aggiunge relazione strutture-dotazioni
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD CONSTRAINT FK_tb_strutture_tb_dichiarazioni " + _
	  "		FOREIGN KEY (online_dichiarazione_id) REFERENCES tb_dichiarazioni(dic_id) " + _
	  "		NOT FOR REPLICATION " + _
	  " ; " + _
	  " ALTER TABLE tb_strutture NOCHECK CONSTRAINT FK_tb_strutture_tb_dichiarazioni "
CALL DB.Execute(sql, 258)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 259
'...........................................................................................
'	upgrade della tabella admin
'...........................................................................................
sql = " ALTER TABLE tb_admin ALTER COLUMN admin_email NVARCHAR(100) NULL;"& _
	  " UPDATE tb_admin SET admin_email = RTRIM(admin_email)"
CALL DB.Execute(sql, 259)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 260
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__5(conn)
CALL DB.Execute(sql, 260)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 261
'...........................................................................................
'	aggiunge campi a tb_loginstru per tracciatura compilazione online
'...........................................................................................
sql = "ALTER TABLE tb_loginstru ADD " + _
	  "		current_online_str_id INT NULL, " + _
	  "		current_online_DataModifica SMALLDATETIME NULL "
CALL DB.Execute(sql, 261)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 262
'...........................................................................................
'	aggiunge campi a tb_loginstru per copmilazione online con termini personalizzati
'...........................................................................................
sql = "ALTER TABLE tb_loginstru ADD " + _
	  "			current_avviso_str_id INT NULL, " + _
	  "			current_avviso_datamodifica SMALLDATETIME NULL, " + _
	  "			custom_dichiarazione BIT NULL, " + _
	  "			custom_dic_tipo INT NULL, " + _
	  "			custom_dic_anno_prezzi INT NULL, " + _
	  "			custom_dic_data_inizio SMALLDATETIME NULL, " + _
	  "			custom_dic_data_fine SMALLDATETIME NULL " + _
	  " ; " + vbCRLf + _
	  " UPDATE tb_loginstru SET custom_dichiarazione=0 ; " + _
	  " ALTER TABLE tb_loginstru ALTER COLUMN custom_dichiarazione BIT NOT NULL ; " + _
	  "ALTER TABLE tb_loginstru DROP COLUMN accesso_personalizzato; " + _
	  "ALTER TABLE tb_loginstru DROP COLUMN compilazione_mod_from; " + _
	  "ALTER TABLE tb_loginstru DROP COLUMN compilazione_mod_to; "
CALL DB.Execute(sql, 262)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 263
'...........................................................................................
'	aggiunge campi a tb_loginstru per gestione dichiarazione ed avviso
'...........................................................................................
sql = " ALTER TABLE tb_loginstru DROP COLUMN current_online_str_id; " + _
	  " ALTER TABLE tb_loginstru DROP COLUMN current_online_DataModifica; " + _
	  " ALTER TABLE tb_loginstru ADD " + _
	  "			current_dich_str_id INT NULL, " + _
	  "			current_dich_datamodifica SMALLDATETIME NULL "
CALL DB.Execute(sql, 263)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 264
'...........................................................................................
'	aggiunge campi a tb_loginstru per gestione dichiarazione ed avviso
'...........................................................................................
sql = " ALTER TABLE tb_loginstru DROP COLUMN current_avviso_str_id; " + _
	  " ALTER TABLE tb_loginstru DROP COLUMN current_avviso_datamodifica; " + _
	  " ALTER TABLE tb_loginstru ADD " + _
	  "			last_dichiarazione_id INT NULL, " + _
	  "			last_dichiarazione_data SMALLDATETIME, " + _
	  "			last_avviso_data SMALLDATETIME "
CALL DB.Execute(sql, 264)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 265
'...........................................................................................
'	aggiunge campi a tb_loginstru per gestione dichiarazione ed avviso
'...........................................................................................
sql = " ALTER TABLE tb_strutture DROP COLUMN online_confermato_open_from; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato_open_to; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato_anno; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato_periodo; " + _
	  " ALTER TABLE tb_strutture ADD " + _
	  "			online_dic_tipo INT NULL, " + _
	  "			online_dic_anno_prezzi INT NULL, " + _
	  "			online_dic_data_inizio SMALLDATETIME, " + _
	  "			online_dic_data_fine SMALLDATETIME; "
CALL DB.Execute(sql, 265)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 266
'...........................................................................................
'	aggiorna vista per gestione login
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		online_dic_presentata BIT NULL, " + _
	  "		online_dic_presentata_data SMALLDATETIME NULL, " + _
	  "		online_dic_presentata_utente nvarchar(50), " + _
      "		online_dic_completata BIT NULL, " + _
	  "		online_dic_completata_data SMALLDATETIME NULL, " + _
	  "		online_dic_completata_utente nvarchar(50) ; " + _
	  " UPDATE tb_strutture SET online_dic_completata=0, online_dic_presentata=0;" + _
	  " ALTER TABLE tb_strutture ALTER COLUMN online_dic_completata BIT NOT NULL ; " + _
	  " ALTER TABLE tb_strutture ALTER COLUMN online_dic_presentata BIT NOT NULL ; " + _
	  " ALTER TABLE tb_loginstru DROP COLUMN last_dichiarazione_id; " + _
	  " ALTER TABLE tb_loginstru DROP COLUMN last_dichiarazione_data; " + _
	  " ALTER TABLE tb_loginstru DROP COLUMN last_avviso_data; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato_data; " + _
	  " ALTER TABLE tb_strutture DROP COLUMN online_confermato_utente; " + _
	  " ALTER TABLE tb_loginstru ADD current_dich_id INT NULL "
CALL DB.Execute(sql, 266)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 267
'...........................................................................................
'	modifica struttura tb_loginstru per mantenimento ultima dichiarazione
'...........................................................................................
sql = " ALTER TABLE tb_loginstru ADD current_dich_tipo INT NULL, " + _
	  "								 current_dich_anno_prezzi INT NULL "
CALL DB.Execute(sql, 267)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 268
'...........................................................................................
'	modifica struttura tb_loginstru per mantenimento ultima dichiarazione
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		online_dic_annullata BIT NULL, " + _
	  "		online_dic_annullata_data SMALLDATETIME NULL, " + _
	  "		online_dic_annullata_utente nvarchar(50) ; " + _
	  " UPDATE tb_strutture SET online_dic_annullata=0;" + _
	  " ALTER TABLE tb_strutture ALTER COLUMN online_dic_annullata BIT NOT NULL ; "
CALL DB.Execute(sql, 268)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 269
'...........................................................................................
'	scorre tutte le strutture per ricalcolare i dati di tb_loginstru
'...........................................................................................
sql = DropObject(conn, "spstr_UPDATE_tb_loginstru", "PROCEDURE") + _
      " CREATE PROCEDURE dbo.spstr_UPDATE_tb_loginstru( " + vbCrLf + _
	  "     @REGCODE nvarchar(12) " + vbCrLf + _
	  " ) " + vbCrLf + _
	  " AS " + vbCrLf + _
	  "     DECLARE @current_str_id INT " + vbCrLf + _
	  "     DECLARE @current_DataModifica SMALLDATETIME " + vbCrLF + _
	  "     DECLARE @current_valid_str_id INT " + vbCrLf + _
	  "     DECLARE @current_valid_DataModifica SMALLDATETIME " + vbCrLF + _
	  "     DECLARE @current_dich_str_id INT " + vbCrLf + _
	  "     DECLARE @current_dich_datamodifica SMALLDATETIME " + vbCrLf + _
	  "     DECLARE @current_dich_id INT " + vbCrLf + _
	  "     DECLARE @current_dich_tipo INT " + vbCrLf + _
	  "     DECLARE @current_dich_anno_prezzi INT " + vbCrLf + _
	  vbcrlf + _
	  "     --recupera dati record corrente " + vbCrLf + _
	  "     SELECT TOP 1 @current_str_id = str_id, " + vbCrLf + _
      "                  @current_DataModifica = DataModifica " + vbCrLf + _
	  "         FROM tb_strutture WHERE RegCode=@REGCODE " + vbCrLf + _
	  "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
	  vbCrLf + _
	  "     --recupera dati ultimo record validato " + vbCrLf + _
	  "     SELECT TOP 1 @current_valid_str_id = str_id, " + vbCrLf + _
      "                  @current_valid_DataModifica = DataModifica " + vbCrLf + _
	  "         FROM tb_strutture WHERE RegCode=@REGCODE AND IsNull(record_validato, 0)=1 " + vbCrLf + _
	  "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
	   vbCrLf + _
	  "     --recupera dati ultimo record completato come dichiarazione " + vbCrLf + _
	  "     SELECT TOP 1 @current_dich_str_id = str_id, " +  vbCrLf + _
      "                  @current_dich_datamodifica = DataModifica, " + vbCrLf + _
	  "                  @current_dich_id = online_dichiarazione_id, " + vbCrLF + _
      "                  @current_dich_tipo = online_dic_tipo, " + vbCrLf + _
	  "                  @current_dich_anno_prezzi = online_dic_anno_prezzi " + vbCrLf + _
	  "         FROM tb_strutture WHERE RegCode=@REGCODE " + vbCrLf + _
      "                             AND IsNull(record_validato, 0)=1 " + vbCrLf + _
      "                             AND IsNull(online_dic_completata, 0)=1 " + vbCrLf+ _
      "                             AND IsNull(online_dic_annullata, 0)= 1 " + vbCrLf + _
	  "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
	  vbCrLf + _
	  "     --aggiorna record tb_loginstru " + vbCrLf + _
	  "     UPDATE tb_loginstru SET " + vbCrLf + _
	  "         current_str_id = @current_str_id, " + vbCRlf + _
	  "         current_DataModifica = @current_DataModifica, " + vbCrlf + _
	  "         current_valid_str_id = @current_valid_str_id, " + vbCrLf + _
	  "         current_valid_DataModifica = @current_valid_DataModifica, " + vbCrLf + _
	  "         current_dich_str_id = @current_dich_str_id, " + vbCrLf + _
	  "         current_dich_datamodifica = @current_dich_datamodifica, " + vbCrLf + _
	  "         current_dich_id = @current_dich_id, " + vbCrLf + _
	  "         current_dich_tipo = @current_dich_tipo, " + vbCrLf + _
	  "         current_dich_anno_prezzi = @current_dich_anno_prezzi " + vbCrLf + _
	  "     WHERE CodAlb = @RegCode " + vbCrLf + _
      " ; " + _
      " DECLARE RS CURSOR " + vbCrLf + _
	  " 	FOR SELECT CodAlb FROM tb_loginstru " + vbCrLf + _
	  vbCrLf + _
	  "	DECLARE @CODALB nvarchar(12) " + vbCrLf+ _
	  vbCrLf + _
	  " OPEN RS " + vbCrLf+ _
	  vbCrLf + _
	  " FETCH NEXT FROM RS INTO @CODALB " + vbCrlf + _
	  "	WHILE (@@fetch_status <> -1) " + vbCrLf + _
	  "		BEGIN " + vbCrLf + _
	  "			IF (@@fetch_status <> -2) " + vbCrLf + _
	  "				BEGIN " + vbCrLf + _
	  "					EXEC spstr_UPDATE_tb_loginstru @CODALB " + vbCrLf + _
	  "				END " + vbCrLf + _
	  " 		FETCH NEXT FROM RS INTO @CODALB " + vbCrLf + _
	  "		END " + vbCrLf + _
	  vbCrLf + _
	  " CLOSE RS " + vbCrLf + _
	  "	DEALLOCATE RS "
CALL DB.Execute(sql, 269)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 270
'...........................................................................................
'	modifica struttura tb_loginstru rimuove campi non piu' utilizzati
'...........................................................................................
sql = " ALTER TABLE tb_loginstru DROP COLUMN compilato_mod_online ; "
CALL DB.Execute(sql, 270)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 271
'...........................................................................................
'	aggiunge colonna per indicazione dell'applicazione che gestisce il modello
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "		mod_applicazione_id INT NULL ; " + _
      " UPDATE tb_modelli SET mod_applicazione_id=206 WHERE mod_id=18 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=207 WHERE mod_id=19 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=209 WHERE mod_id=20 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=210 WHERE mod_id=21 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=210 WHERE mod_id=22 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=211 WHERE mod_id=23 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=212 WHERE mod_id=24 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=215 WHERE mod_id=25 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=216 WHERE mod_id=26 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=217 WHERE mod_id=27 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=218 WHERE mod_id=28 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=219 WHERE mod_id=29 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=220 WHERE mod_id=30 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=220 WHERE mod_id=31 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=221 WHERE mod_id=32 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=222 WHERE mod_id=33 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=223 WHERE mod_id=34 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=223 WHERE mod_id=35 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=223 WHERE mod_id=36 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=225 WHERE mod_id=37 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=226 WHERE mod_id=38 ;" + _
      " UPDATE tb_modelli SET mod_applicazione_id=227 WHERE mod_id=39 ;" + _
	  " ALTER TABLE tb_modelli ALTER COLUMN " + _
	  "		mod_applicazione_id INT NOT NULL ; " + _
	  " ALTER TABLE tb_modelli ADD CONSTRAINT FK_tb_modelli_tb_siti " + _
	  "		FOREIGN KEY (mod_applicazione_id) REFERENCES tb_siti(id_sito) " + _
	  "		ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
	  " ALTER TABLE tb_modelli DROP CONSTRAINT DF__tb_modell__Mod_C__5A846E65; " + _
	  " ALTER TABLE tb_modelli DROP CONSTRAINT DF__tb_modell__Mod_F__6501FCD8; "
CALL DB.Execute(sql, 271)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 272
'...........................................................................................
'	rimuove relazione tra strutture e dichiarazioni non applicabile.
'...........................................................................................
sql = " ALTER TABLE tb_strutture DROP CONSTRAINT FK_tb_strutture_tb_dichiarazioni ; "
CALL DB.Execute(sql, 272)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 273
'...........................................................................................
'	ingrandisce campi testo per dichiarazioni e servizi
'...........................................................................................
sql = " ALTER TABLE rel_str_serv ALTER COLUMN rel_str_serv_val nvarchar(500) NULL ; " + _
	  " ALTER TABLE rel_str_dotaz ALTER COLUMN rel_str_dotaz_testo_it nvarchar(500) NULL ; "
CALL DB.Execute(sql, 273)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 274
'...........................................................................................
'	aggiunge relazione per gestione permessi aggiuntivi applicativi interni
'...........................................................................................
sql = " CREATE TABLE dbo.tb_turismo_admin_sito ( " + _
	  "		tas_id INT IDENTITY (1, 1) NOT NULL , " + _
	  "		tas_admin_id INT NOT NULL , " + _
	  "		tas_sito_id INT NOT NULL , " + _
	  "		tas_apt NVARCHAR(2) NULL , " + _
	  "		tas_permesso INT NOT NULL " + _
	  " ); " + _
	  " ALTER TABLE tb_turismo_admin_sito ADD CONSTRAINT PK_tb_turismo_admin_sito PRIMARY KEY NONCLUSTERED (tas_id); "+ _
	  " ALTER TABLE tb_turismo_admin_sito ADD"+ _
	  " 	CONSTRAINT FK_tb_turismo_admin_sito_tb_siti FOREIGN KEY (tas_sito_id) REFERENCES tb_siti(id_sito)"+ _
	  "		ON UPDATE CASCADE ON DELETE CASCADE," + _
	  " 	CONSTRAINT FK_tb_turismo_admin_sito_tb_admin FOREIGN KEY (tas_admin_id) REFERENCES tb_admin(id_admin)"+ _
	  "		ON UPDATE CASCADE ON DELETE CASCADE, " + _
	  " 	CONSTRAINT FK_tb_turismo_admin_sito_tb_apt FOREIGN KEY (tas_apt) REFERENCES tb_apt(apt_codice)"+ _
	  "		ON UPDATE NO ACTION ON DELETE NO ACTION;"+ _
	  " ALTER TABLE tb_turismo_admin_sito NOCHECK CONSTRAINT FK_tb_turismo_admin_sito_tb_apt;"
CALL DB.Execute(sql, 274)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 275
'...........................................................................................
'	riempie dati per utenti degli applicativi di gestione dati strutture
'...........................................................................................
sql = " INSERT INTO tb_turismo_admin_sito (tas_admin_id, tas_sito_id, tas_apt, tas_permesso) " + _
	  "		SELECT tb_turismo_admin.id_admin, id_sito, " + _
	  "			   CASE WHEN admin_apt = '' THEN NULL ELSE admin_apt END, " + _
	  "			   CASE (SELECT MIN(rel_as_permesso) FROM rel_admin_sito WHERE admin_id = tb_admin.id_admin AND sito_id = tb_siti.id_sito) " + _
	  "					WHEN 1 THEN " & P_ADMINISTRATOR & _
	  "					WHEN 2 THEN " & P_VALIDATOR & _
	  "					WHEN 3 THEN " & P_READ_ONLY & _
	  "					WHEN 4 THEN " & P_MOD_IDENTITY & _
	  "					ELSE " & P_READ_ONLY & _
	  "			   END " + _
	  "		FROM (tb_admin INNER JOIN tb_turismo_admin ON tb_admin.id_admin = tb_turismo_admin.id_admin), tb_siti " + _
	  "		WHERE tb_siti.id_sito IN (SELECT mod_applicazione_id FROM tb_modelli) " + _
	  "			  AND tb_admin.id_admin IN (SELECT admin_id FROM rel_admin_sito " + _
	  "			  							INNER JOIN tb_modelli ON rel_admin_sito.sito_id = tb_modelli.mod_applicazione_id) " + _
	  "			  AND EXISTS(SELECT 1 FROM rel_admin_sito WHERE admin_id = tb_admin.id_admin AND sito_id = tb_siti.id_sito) " + _
	  "	; " + _
	  " DELETE FROM rel_admin_sito WHERE admin_id IN (SELECT tas_admin_id FROM tb_turismo_admin_sito) " + _
	  "		AND sito_id IN (SELECT mod_applicazione_id FROM tb_modelli) " + _
	  " ; " + _
	  "	INSERT INTO rel_admin_sito (admin_id, sito_id, rel_as_permesso ) " + _
	  "		SELECT tas_admin_id, tas_sito_id, 1 FROM tb_turismo_admin_sito " + _
	  " ; " + _
	  " UPDATE tb_siti SET " + _
	  "			sito_p1 = REPLACE(sito_p1, '_ADMIN', '_USER'), " + _
	  "			sito_p2 = '', " + _
	  "			sito_p3 = '', " + _
  	  "			sito_p4 = '', " + _
  	  "			sito_p5 = '', " + _
  	  "			sito_p6 = '', " + _
  	  "			sito_p7 = '', " + _
  	  "			sito_p8 = '', " + _
  	  "			sito_p9 = '' " + _
	  "		WHERE id_sito IN (SELECT mod_applicazione_id FROM tb_modelli)" + _
	  " ; " + _
	  DropObject(conn, "tb_turismo_admin", "TABLE") 
CALL DB.Execute(sql, 275)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 276
'...........................................................................................
' 	inserisce dati dichiarazioni 
'	imposta vecchie dichiarazioni per le strutture che hanno compilato online
'...........................................................................................
sql = " SET IDENTITY_INSERT tb_dichiarazioni ON" + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (1, 2004,  CONVERT(DATETIME, '2003-07-14 00:00:00', 102),  CONVERT(DATETIME, '2003-10-02 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (2, 2004,  CONVERT(DATETIME, '2004-02-01 00:00:00', 102),  CONVERT(DATETIME, '2004-03-01 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (3, 2005,  CONVERT(DATETIME, '2004-07-14 00:00:00', 102),  CONVERT(DATETIME, '2004-10-01 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (4, 2005,  CONVERT(DATETIME, '2005-02-01 00:00:00', 102),  CONVERT(DATETIME, '2005-02-28 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (5, 2006,  CONVERT(DATETIME, '2005-07-14 00:00:00', 102),  CONVERT(DATETIME, '2005-10-03 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (6, 2006,  CONVERT(DATETIME, '2006-02-01 00:00:00', 102),  CONVERT(DATETIME, '2006-03-01 00:00:00', 102)) " + _
	  " INSERT INTO tb_dichiarazioni(dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine) VALUES (7, 2007,  CONVERT(DATETIME, '2006-06-01 00:00:00', 102),  CONVERT(DATETIME, '2006-10-01 00:00:00', 102)) " + _
	  " SET IDENTITY_INSERT tb_dichiarazioni OFF" + _
	  " " + _
	  " SET IDENTITY_INSERT rel_dichiarazioni_modelli ON" + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (1, 1, 18, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (2, 1, 19, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (3, 1, 20, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (4, 1, 22, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (5, 1, 23, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (6, 1, 24, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (7, 1, 25, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (8, 1, 26, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (9, 1, 27, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (10, 1, 28, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (11, 1, 29, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (12, 1, 30, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (13, 1, 32, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (14, 1, 33, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (15, 2, 18, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (16, 2, 19, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (17, 2, 20, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (18, 2, 22, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (19, 2, 27, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (20, 2, 28, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (21, 3, 18, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (22, 3, 19, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (23, 3, 20, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (24, 3, 22, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (25, 3, 23, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (26, 3, 24, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (27, 3, 25, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (28, 3, 26, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (29, 3, 27, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (30, 3, 28, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (31, 3, 29, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (32, 3, 30, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (33, 3, 32, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (34, 3, 33, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (35, 3, 34, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (36, 4, 18, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (37, 4, 19, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (38, 4, 20, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (39, 4, 22, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (40, 4, 27, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (41, 4, 28, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (42, 5, 18, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (43, 5, 19, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (44, 5, 20, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (45, 5, 22, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (46, 5, 23, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (47, 5, 24, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (48, 5, 25, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (49, 5, 26, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (50, 5, 27, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (51, 5, 28, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (52, 5, 29, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (53, 5, 30, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (54, 5, 32, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (55, 5, 33, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (56, 5, 34, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (57, 6, 18, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (58, 6, 19, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (59, 6, 20, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (60, 6, 22, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (61, 6, 27, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (62, 6, 28, 2) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (63, 7, 18, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (64, 7, 19, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (65, 7, 20, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (66, 7, 22, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (67, 7, 23, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (68, 7, 24, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (69, 7, 25, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (70, 7, 26, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (71, 7, 27, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (72, 7, 28, 1) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (73, 7, 29, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (74, 7, 30, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (75, 7, 32, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (76, 7, 33, 0) " + _
	  " INSERT INTO rel_dichiarazioni_modelli(rel_id, rel_dic_id, rel_mod_id, rel_tipo_dichiarazione) VALUES (77, 7, 34, 0) " + _
	  " SET IDENTITY_INSERT rel_dichiarazioni_modelli OFF" + _
	  " ;" + _
	  " DECLARE RS CURSOR " + vbCrLf + _ 
	  "	FOR SELECT RegCode, MAX(str_id), " + _
	  "			(CONVERT(DATETIME, str(YEAR(MAX(tb_str_logs.str_log_data))) + '-' + str(MONTH(MAX(tb_str_logs.str_log_data))) + '-' + str(DAY(MAX(tb_str_logs.str_log_data))) + ' 00:00:00', 102)) AS DataModifica,  " + _
	  "			dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine, rel_tipo_dichiarazione, " + _
	  "			(LTRIM(RIGHT(MAX(str_log_ope),  LEN(MAX(str_log_ope)) - CHARINDEX(':', MAX(str_log_ope))))) AS utente " + _
	  "		FROM tb_str_logs INNER JOIN tb_strutture ON tb_str_logs.str_log_record = tb_strutture.str_id " + _
	  "			INNER JOIN tb_loginstru ON tb_strutture.RegCode = tb_loginstru.codalb " + _
	  "			INNER JOIN tb_dichiarazioni ON CONVERT(DATETIME, str(YEAR(tb_str_logs.str_log_data)) + '-' + str(MONTH(tb_str_logs.str_log_data)) + '-' + str(DAY(tb_str_logs.str_log_data)) + ' 00:00:00', 102) " + _
	  "										   BETWEEN tb_dichiarazioni.dic_data_inizio AND tb_dichiarazioni.dic_data_fine " + _
	  "			INNER JOIN rel_dichiarazioni_modelli ON tb_dichiarazioni.dic_id = rel_dichiarazioni_modelli.rel_dic_id " + _
	  "													AND rel_dichiarazioni_modelli.rel_mod_id=tb_loginstru.modello " + _
	  "		WHERE str_log_des LIKE '%modulo%' AND str_log_des LIKE '%corretta%' " + _
	  "		GROUP BY RegCode, dic_id, dic_anno_prezzi, dic_data_inizio, dic_data_fine, rel_tipo_dichiarazione " + _
	  "		ORDER BY DataModifica " + _
	  vbCrLf + _
	  "	DECLARE @RegCode nvarchar(12) " + vbCrLf + _
	  "	DECLARE @str_id INT " + vbCrLF + _
	  "	DECLARE @DataModifica SMALLDATETIME " + vbCrLf + _
	  "	DECLARE @dic_id INT " + vbCrLf + _
	  "	DECLARE @dic_anno_prezzi INT " + vbCrLf + _
	  "	DECLARE @dic_data_inizio SMALLDATETIME " + vbCrLf + _
	  "	DECLARE @dic_data_fine SMALLDATETIME " + vbCrLF + _
	  "	DECLARE @rel_tipo_dichiarazione INT " + vbCrLF + _
	  "	DECLARE @utente nvarchar(50) " + vbCrLF + _
	  vbCrLF + _
	  "	OPEN RS " + vbCrLF + _
	  vbCrLF + _
	  "	FETCH NEXT FROM RS INTO @RegCode, @Str_id, @DataModifica, @dic_id, @dic_anno_prezzi, @dic_data_inizio, @Dic_data_fine, @rel_tipo_dichiarazione, @utente " + vbCRLf + _
	  "	WHILE (@@fetch_status <> -1) " + vbCrLf + _
	  "		BEGIN " + vbCrLf + _
	  "			UPDATE tb_strutture SET " + _
	  "				UtenteModifica = @Utente, " + _
	  "				online_modifica_data = @DataModifica, " + _
	  "				online_modifica_utente = @Utente, " + _
	  "				online_dichiarazione_id = @Dic_id, " + _
	  "				online_dic_tipo = @rel_tipo_dichiarazione, " + _
	  "				online_dic_anno_prezzi = @dic_anno_prezzi, " + _
	  "				online_dic_data_inizio = @dic_data_inizio, " + _
	  "				online_dic_data_fine = @dic_data_fine, " + _
	  "				online_dic_presentata = 1, " + _
	  "				online_dic_presentata_data = @DataModifica, " + _
	  "				online_dic_presentata_utente = @Utente, " + _
	  "				online_dic_completata = 1, " + _
	  "				online_dic_completata_data = @dataModifica, " + _
	  "				online_dic_completata_utente = @utente " + _
	  "			WHERE str_id = @str_id AND RegCode = @RegCode " + vbCrLF + _
	  "			FETCH NEXT FROM RS INTO @RegCode, @Str_id, @DataModifica, @dic_id, @dic_anno_prezzi, @dic_data_inizio, @Dic_data_fine, @rel_tipo_dichiarazione, @utente " + vbCRLf + _
	  "		END " + vbcrLf + _
	  vbCrLf+ _
	  "	CLOSE RS " + vbCrLf + _
	  " DEALLOCATE RS " + vbCrLF + _
	  vbCrLF + _
	  " DECLARE RSD CURSOR " + vbCrLF + _
	  "		FOR SELECT RegCode FROM tb_strutture WHERE IsNull(online_dichiarazione_id, 0)>0 GROUP BY Regcode " + vbCrLF + _
	  vbCrLF + _
	  "	OPEN RSD " + vbCrLF + _
	  vbCrLF + _
	  "	FETCH NEXT FROM RSD INTO @RegCode " + vbCrLf + _
	  "	WHILE (@@fetch_status <> -1) " + vbCrLf + _
	  "		BEGIN " + vbCrLf + _
	  "			EXEC spstr_UPDATE_tb_loginstru @RegCode " + vbCrLf + _
	  "			FETCH NEXT FROM RSD INTO @RegCode " + vbCrLf + _
	  "		END " + vbcrLf + _
	  vbCrLf+ _
	  "	CLOSE RSD " + vbCrLf + _
	  " DEALLOCATE RSD " + vbCrLF
CALL DB.Execute(sql, 276)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 277
'...........................................................................................
'	ingrandisce campi testo per dichiarazioni e servizi
'...........................................................................................
sql = " CREATE TABLE dbo.tb_modelli_default ( " + _
	  "		default_anno_prezzi INT NOT NULL, " +_
	  "		default_area_online_disabilitata BIT NOT NULL " + _
	  "	) " + _
	  " ; " + _
	  " INSERT INTO tb_modelli_default(default_anno_prezzi, default_area_online_disabilitata) " + _
	  "		VALUES 					  (2007               , 0                               ) "
CALL DB.Execute(sql, 277)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 278
'...........................................................................................
'	ingrandisce campi testo per dichiarazioni e servizi
'...........................................................................................
sql = DropObject(conn, "tr__tb_admin__insert", "TRIGGER") + _
	  DropObject(conn, "tr__tb_admin__delete", "TRIGGER")
CALL DB.Execute(sql, 278)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 279
'...........................................................................................
'	crea le tabella per l'applicativo delle statistiche
'...........................................................................................
sql = " CREATE TABLE dbo.tb_turismo_statTabelleTemp ("+ _
	  " 	stt_nome nvarchar (150) NOT NULL ,"+ _
	  "		stt_data smalldatetime NULL ,"+ _
	  "		CONSTRAINT PK_tb_turismo_statTabelleTemp PRIMARY KEY CLUSTERED (stt_nome)"+ _
	  "	);"+ _
	  " CREATE TABLE dbo.tb_turismo_statCampi ("+ _
	  "		stc_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		stc_campo nvarchar (250) NOT NULL ,"+ _
	  "		stc_tipo int NOT NULL ,"+ _
	  "		stc_tabella nvarchar (250) NOT NULL ,"+ _
	  "		stc_pos int NULL ,"+ _
	  "		CONSTRAINT PK_tb_turismo_statCampi PRIMARY KEY CLUSTERED (stc_id)"+ _
	  "	);"+ _
	  " CREATE TABLE dbo.rel_turismo_statCampiModelli ("+ _
	  "		rcm_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		rcm_gruppo_id int NULL ,"+ _
	  "		rcm_modello_id int NULL ,"+ _
	  "		rcm_campo_id int NULL ,"+ _
	  "		rcm_descr nvarchar (250) NULL ,"+ _
	  "		rcm_protetto bit NULL ,"+ _
	  "		rcm_valID int NULL ,"+ _
	  "		rcm_sorgente nvarchar (250) NULL ,"+ _
	  "		rcm_paginaGestione nvarchar (250) NULL ,"+ _
	  "		CONSTRAINT PK_rel_turismo_campiModelli PRIMARY KEY CLUSTERED (rcm_id),"+ _
	  "		CONSTRAINT FK_rel_turismo_statCampiModelli_tb_grp_vis"+ _
	  "			FOREIGN KEY (rcm_gruppo_id) REFERENCES tb_grp_vis (Grp_id),"+ _
	  "		CONSTRAINT FK_rel_turismo_statCampiModelli_tb_turismo_statCampi"+ _
	  "			FOREIGN KEY (rcm_campo_id) REFERENCES tb_turismo_statCampi (stc_id) ON DELETE CASCADE  ON UPDATE CASCADE"+ _
	  "	);"+ _
	  " alter table dbo.rel_turismo_statCampiModelli nocheck constraint FK_rel_turismo_statCampiModelli_tb_grp_vis;"+ _
	  " CREATE TABLE dbo.tb_turismo_statSet ("+ _
	  "		sts_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		sts_nome nvarchar (100) NOT NULL ,"+ _
	  "		sts_descr ntext NULL ,"+ _
	  "		CONSTRAINT PK_tb_turismo_statSet PRIMARY KEY CLUSTERED (sts_id)"+ _
	  "	);"+ _
	  " CREATE TABLE dbo.rel_turismo_statSetCampi ("+ _
	  "		rsc_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		rsc_campo_id int NOT NULL ,"+ _
	  "		rsc_set_id int NOT NULL ,"+ _
	  "		CONSTRAINT PK_rel_turismo_statSetCampi PRIMARY KEY CLUSTERED (rsc_id),"+ _
	  "		CONSTRAINT FK_rel_turismo_statSetCampi_rel_turismo_statCampiModelli"+ _
	  "			FOREIGN KEY (rsc_campo_id) REFERENCES rel_turismo_statCampiModelli (rcm_id) ON DELETE CASCADE ON UPDATE CASCADE,"+ _
	  "		CONSTRAINT FK_rel_turismo_statSetCampi_tb_turismo_statSet"+ _
	  "			FOREIGN KEY (rsc_set_id) REFERENCES tb_turismo_statSet (sts_id) ON DELETE CASCADE  ON UPDATE CASCADE"+ _
	  "	);"+ _
	  "	CREATE TABLE dbo.tb_turismo_statCriteri ("+ _
	  "		scr_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		scr_set_id int NULL ,"+ _
	  "		scr_nome nvarchar (100) NOT NULL ,"+ _
	  "		scr_descr ntext NULL ,"+ _
	  "		scr_dataC datetime NOT NULL ,"+ _
	  "		scr_temp bit NULL ,"+ _
	  "		CONSTRAINT PK_tb_turismo_statCriteri PRIMARY KEY CLUSTERED (scr_id),"+ _
	  "		CONSTRAINT FK_tb_turismo_statCriteri_tb_turismo_statSet"+ _
	  "			FOREIGN KEY (scr_set_id) REFERENCES tb_turismo_statSet (sts_id)"+ _
	  "	);"+ _
	  "	alter table dbo.tb_turismo_statCriteri nocheck constraint FK_tb_turismo_statCriteri_tb_turismo_statSet;"+ _
	  "	CREATE TABLE dbo.rel_turismo_statCriteriCampi ("+ _
	  "		rcc_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		rcc_criterio_id int NOT NULL ,"+ _
	  "		rcc_campo_id int NOT NULL ,"+ _
	  "		rcc_valore nvarchar (250) NOT NULL ,"+ _
	  "		rcc_confronto int NULL ,"+ _
	  "		CONSTRAINT PK_rel_turismo_statCriteriCampi PRIMARY KEY CLUSTERED (rcc_id),"+ _
	  "		CONSTRAINT FK_rel_turismo_statCriteriCampi_rel_turismo_statCampiModelli"+ _
	  "	 		FOREIGN KEY (rcc_campo_id) REFERENCES rel_turismo_statCampiModelli (rcm_id) ON DELETE CASCADE ON UPDATE CASCADE,"+ _
	  "		CONSTRAINT FK_rel_turismo_statCriteriCampi_tb_turismo_statCriteri"+ _
	  "			FOREIGN KEY (rcc_criterio_id) REFERENCES tb_turismo_statCriteri (scr_id) ON DELETE CASCADE ON UPDATE CASCADE"+ _
	  "	);"+ _
	  "	CREATE TABLE dbo.rel_turismo_statOrdineCampi ("+ _
	  "		roc_id int IDENTITY (1, 1) NOT NULL ,"+ _
	  "		roc_criterio_id int NOT NULL ,"+ _
	  "		roc_campo_id int NOT NULL ,"+ _
	  "		roc_ordine int NOT NULL ,"+ _
	  "		roc_tipo bit NULL ,"+ _
	  "		CONSTRAINT PK_rel_turismo_statOrdineCampi PRIMARY KEY CLUSTERED (roc_id),"+ _
	  "		CONSTRAINT FK_rel_turismo_statOrdineCampi_rel_turismo_statCampiModelli"+ _
	  "			FOREIGN KEY (roc_campo_id) REFERENCES rel_turismo_statCampiModelli (rcm_id) ON DELETE CASCADE ON UPDATE CASCADE,"+ _
	  "		CONSTRAINT FK_rel_turismo_statOrdineCampi_tb_turismo_statCriteri1"+ _
	  "			FOREIGN KEY (roc_criterio_id) REFERENCES tb_turismo_statCriteri (scr_id) ON DELETE CASCADE ON UPDATE CASCADE"+ _
	  "	);"
CALL DB.Execute(sql, 279)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 280
'...........................................................................................
'	statistiche - aggiunge flag privati a set dati
'...........................................................................................
sql = " ALTER TABLE tb_turismo_statSet ADD"+ _
	  " 	sts_protetto BIT NULL"
CALL DB.Execute(sql, 280)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 281
'...........................................................................................
'	aggiornamento procedure per webservice
'...........................................................................................
sql = DropObject(conn, "spWS_Hotel_Dotazioni", "PROCEDURE") + _
      DropObject(conn, "spWS_Hotel_Header", "PROCEDURE") + _
      DropObject(conn, "spWS_Hotel_List", "PROCEDURE") + _
      DropObject(conn, "spWS_Hotel_Servizi", "PROCEDURE") + _
      DropObject(conn, "spWS_Hotel_Updated", "PROCEDURE") + _
      DropObject(conn, "spWS_Last_Update", "PROCEDURE") + _
	  " CREATE  PROCEDURE dbo.spWS_Hotel_Dotazioni " + vbCrLF + _
	  "     ( @REGCODE VARCHAR(12)) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     DECLARE @STR_ID int		--id della struttura o dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @PRO_ID int		--id del proprietario dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @TYP_ID int		--id della tipologia dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLF + _
	  vbCrLf + _
	  "     SELECT @STR_ID = Current_valid_Str_id, @TIPO_RECORD=mod_tipo_record " + vbCrLF + _
	  "         FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginStru.Modello=tb_modelli.mod_id " + vbCrLF + _
	  "         WHERE CodAlb = @REGCODE " + vbCrLF + _
	  vbCrLF + _
	  "     IF (@TIPO_RECORD='U') " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             DECLARE @PRO_REGCODE nvarchar(12) " + vbCrLF + _
	  "             DECLARE @TYP_REGCODE nvarchar(12) " + vbCrLF + _
	  "             SELECT @PRO_REGCODE=Cod_Proprietario, @TYP_REGCODE=Cod_Tipologia " + vbCrLF + _
	  "                 FROM tb_Strutture " + vbCrLF + _
	  "                 WHERE STR_ID=@STR_ID " + vbCrLF + _
	  "             SELECT @PRO_ID = Current_valid_str_id " + vbCrLF + _
	  "                 FROM tb_LoginStru " + vbCrLF + _
	  "                 WHERE CodAlb = @PRO_REGCODE " + vbCrLF + _
	  "             IF (@TYP_REGCODE IS NULL) " + vbCrLF + _
	  "                 SET @TYP_ID = 0 " + vbCrLF + _
	  "             ELSE " + vbCrLF + _
	  "                 SELECT @TYP_ID = Current_valid_str_id " + vbCrLF + _
	  "                     FROM tb_LoginStru " + vbCrLF + _
	  "                     WHERE CodAlb = @TYP_REGCODE " + vbCrLF + _
	  "         END " + vbCrLF + _
	  "     ELSE " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             SET @PRO_ID = 0 " + vbCrLF + _
	  "             SET @TYP_ID = 0 " + vbCrLF + _
	  "         END " + vbCrLF + _
	  vbCrLf + _
	  "     SELECT (tb_pubblicazioni_APT.pub_label_it) AS gruppo_it, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_en) AS gruppo_en, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_fr) AS gruppo_fr, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_de) AS gruppo_de, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_es) AS gruppo_es, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_APT_nome_ITA) AS nome_it, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_APT_nome_ENG) AS nome_en, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_APT_nome_FRA) AS nome_fr, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_APT_nome_TED) AS nome_de, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_APT_nome_SPA) AS nome_es, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_symb) AS simbolo, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_typ) AS tipo, " + vbCrLF + _
	  "         (tb_dotazioni.dotaz_num_val) AS numero_valori, " + vbCrLF + _
	  "         (dotaz_lbl_1_1) AS label_1_level_1, " + vbCrLF + _
	  "         (dotaz_lbl_1_2) AS label_2_level_1, " + vbCrLF + _
	  "         (dotaz_lbl_1_3) AS label_3_level_1, " + vbCrLF + _
	  "         (dotaz_lbl_1_4) AS label_4_level_1, " + vbCrLF + _
	  "         (dotaz_lbl_2_1) AS label_1_level_2, " + vbCrLF + _
	  "         (dotaz_lbl_2_2) AS label_2_level_2, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' OR dotaz_typ='MN' OR dotaz_typ='MP' THEN rel_1.rel_str_dotaz_testo_it ELSE CAST(rel_1.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore1, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' OR dotaz_typ='FM' THEN rel_2.rel_str_dotaz_testo_it ELSE CAST(rel_2.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore2, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_3.rel_str_dotaz_testo_it ELSE CAST(rel_3.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore3, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_4.rel_str_dotaz_testo_it ELSE CAST(rel_4.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore4, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_5.rel_str_dotaz_testo_it ELSE CAST(rel_5.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore5, " + vbCrLF + _
	  "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_6.rel_str_dotaz_testo_it ELSE CAST(rel_6.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore6 " + vbCrLF + _
	  "     FROM tb_dotazioni  " + vbCrLF + _
	  "         INNER JOIN rel_grp_dotaz ON tb_dotazioni.dotaz_id=rel_grp_dotaz.rel_grp_id_dotaz " + vbCrLF + _
	  "         INNER JOIN tb_pubblicazioni_APT ON dotaz_APT_pubblicazione = tb_pubblicazioni_APT.pub_id " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_1 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_1.rel_str_id_dotaz AND rel_1.rel_str_dotaz_pos_val=1 AND (rel_1.rel_id_str_dotaz=@STR_ID OR rel_1.rel_id_str_dotaz=@PRO_ID OR rel_1.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_2 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_2.rel_str_id_dotaz AND rel_2.rel_str_dotaz_pos_val=2 AND (rel_2.rel_id_str_dotaz=@STR_ID OR rel_2.rel_id_str_dotaz=@PRO_ID OR rel_2.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_3 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_3.rel_str_id_dotaz AND rel_3.rel_str_dotaz_pos_val=3 AND (rel_3.rel_id_str_dotaz=@STR_ID OR rel_3.rel_id_str_dotaz=@PRO_ID OR rel_3.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_4 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_4.rel_str_id_dotaz AND rel_4.rel_str_dotaz_pos_val=4 AND (rel_4.rel_id_str_dotaz=@STR_ID OR rel_4.rel_id_str_dotaz=@PRO_ID OR rel_4.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_5 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_5.rel_str_id_dotaz AND rel_5.rel_str_dotaz_pos_val=5 AND (rel_5.rel_id_str_dotaz=@STR_ID OR rel_5.rel_id_str_dotaz=@PRO_ID OR rel_5.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "         LEFT JOIN rel_str_dotaz rel_6 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_6.rel_str_id_dotaz AND rel_6.rel_str_dotaz_pos_val=6 AND (rel_6.rel_id_str_dotaz=@STR_ID OR rel_6.rel_id_str_dotaz=@PRO_ID OR rel_6.rel_id_str_dotaz=@TYP_ID) " + vbCrLF + _
	  "     WHERE tb_pubblicazioni_APT.pub_order>0  " + vbCrLF + _
	  "         AND	rel_grp_dotaz_id IN (SELECT rel_str_id_dotaz FROM rel_str_dotaz WHERE (rel_id_str_dotaz=@STR_ID OR rel_id_str_dotaz=@PRO_ID OR rel_id_str_dotaz=@TYP_ID) AND ((rel_str_dotaz_valore <> 0 AND rel_Str_dotaz_valore IS NOT NULL) OR (rel_str_dotaz_testo_it <> '' AND rel_Str_dotaz_testo_it IS NOT NULL))) " + vbCrLF + _
	  "     ORDER BY tb_pubblicazioni_APT.pub_order, rel_grp_dotaz.dotaz_APT_ordine ,tb_dotazioni.dotaz_APT_nome_ita  " + vbCrLF + _
	  " ; " + _
	  " CREATE   PROCEDURE dbo.spWS_Hotel_Header(@REGCODE nvarchar(12)) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLF + _
	  "     SELECT @TIPO_RECORD=mod_tipo_record " + vbCrLF + _
	  "         FROM tb_modelli INNER JOIN tb_loginStru ON tb_modelli.mod_id=tb_loginstru.modello" + vbCrLF + _
	  "     WHERE CODALB=@REGCODE " + vbCrLF + _
	  vbCrLf + _
	  "     IF (@TIPO_RECORD LIKE '%U%') " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             --unit abitative " + vbCrLF + _
	  "             SELECT UA.regcode, " + vbCrLF + _
	  "                 (SELECT MAX(tb_loginstru.current_valid_DataModifica) " + vbCrLF + _
	  "                     FROM tb_loginstru " + vbCrLF + _
	  "                     WHERE tb_loginstru.CodAlb = UA.RegCode OR " + vbCrLF + _
	  "                         tb_loginstru.CodAlb = UA.Cod_Proprietario OR " + vbCrLF + _
	  "                         tb_loginstru.CodAlb = UA.Cod_Tipologia) AS datamodifica, " + vbCrLF + _
	  "                 UA.tipo, UA.categoria, UA.indirizzo, " + vbCrLF + _
	  "                 UA.civico, UA.localita, UA.cap, UA.provincia, " + vbCrLF + _
	  "                 PRO.telefono, PRO.fax, PRO.email, PRO.weburl, UA.comune, " + vbCrLF + _
	  "                 (CASE WHEN ISNULL(UA.Denominazione,'') =''THEN PRO.Denominazione WHEN PRO.mod_tipo_record='A' THEN (UA.Denominazione + ' - ' + PRO.Denominazione) ELSE UA.Denominazione END) AS denominazione, " + vbCrLF + _
	  "                 (UA.comuneTXT) AS comune_txt , " + vbCrLF + _
	  "                 (UA.comune + '-' + (SELECT Loc_Cod FROM tb_localita " + vbCrLF + _
	  "                                         WHERE REPLACE(Loc_Nome, ' ', '') LIKE REPLACE(UA.localita, ' ', '') " + vbCrLF + _
	  "                                               AND Loc_comune = UA.Comune)) AS loc_cod, " + vbCrLF + _
	  "                 (CASE WHEN (UA.F_REVOCA_LIC='S' OR UA.F_REVOCA_CL='S' OR UA.F_RIM_VINC='S') THEN 1 ELSE 0 END) AS cessataatt, " + vbCrLF + _
	  "                 UA.mod_tipo_classificazione " + vbCrLF + _
	  "             FROM VIEW_valid_Strutture UA INNER JOIN VIEW_valid_strutture PRO ON UA.Cod_proprietario = PRO.RegCode " + vbCrLF + _
	  "             WHERE UA.RegCode = @REGCODE " + vbCrLF + _
	  "         END " + vbCrLF + _
	  "     ELSE " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             --altri tipi di record " + vbCrLF + _
	  "             SELECT regcode, datamodifica, tipo, categoria, indirizzo, civico, localita, cap, provincia, " + vbCrLF + _
	  "                 telefono, fax, email, weburl, comune, " + vbCrLF + _
	  "                 (CASE WHEN ((tipoImmobile LIKE 'D') AND (CHARINDEX('DIP',UPPER(Denominazione))<1)) " + vbCrLF + _
	  "                     THEN RTRIM(Denominazione) + ' DIPENDENZA' " + vbCrLF + _
	  "                     ELSE Denominazione END) AS denominazione, " + vbCrLF + _
	  "                 (comuneTXT) AS comune_txt , " + vbCrLF + _
	  "                 (comune + '-' + (SELECT Loc_Cod FROM tb_localita " + vbCrLF + _
	  "                                  WHERE REPLACE(Loc_Nome, ' ', '') LIKE REPLACE(localita, ' ', '') " + vbCrLF + _
	  "                                        AND Loc_comune = Comune)) AS loc_cod, " + vbCrLF + _
	  "                 (CASE WHEN (F_REVOCA_LIC='S' OR F_REVOCA_CL='S' OR F_RIM_VINC='S') THEN 1 ELSE 0 END) AS cessataatt, " + vbCrLF + _
	  "                 mod_tipo_classificazione " + vbCrLF + _
	  "             FROM VIEW_valid_Strutture " + vbCrLF + _
	  "             WHERE RegCode = @REGCODE " + vbCrLF + _
	  "         END " + vbCrLF + _
	  " ; " + _
	  " CREATE PROCEDURE dbo.spWS_Hotel_List(@APT nvarchar(2)) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     --modello 36 ==> modello unit abitative gestite da agenzie immobiliari " + vbCrLF + _
	  "     SELECT regcode, " + vbCrLF + _
	  "         (CASE WHEN mod_tipo_record='S' OR mod_tipo_record='A' THEN tb_strutture.Datamodifica " + vbCrLF + _
	  "               ELSE (SELECT MAX(current_valid_DataModifica) " + vbCrLF + _
	  "                         FROM tb_loginstru " + vbCrLF + _
	  "                         WHERE tb_loginstru.CodAlb = tb_strutture.RegCode OR " + vbCrLF + _
	  "                               tb_loginstru.CodAlb = tb_strutture.Cod_Proprietario OR " + vbCrLF + _
	  "                               tb_loginstru.CodAlb = tb_strutture.Cod_Tipologia) END) AS datamodifica, " + vbCrLF + _
	  "         (CASE WHEN (mod_tipo_record='U' AND ISNULL(Denominazione,'')='') " + vbCrLF + _
	  "               THEN (SELECT AGE.Denominazione FROM VIEW_valid_testata_Strutture AGE " + vbCrLF + _
	  "                         WHERE AGE.RegCode = tb_strutture.Cod_Proprietario) " + vbCrLF + _
	  "               WHEN (modello=36) " + vbCrLF + _
	  "               THEN (SELECT tb_Strutture.Denominazione + ' - ' + AGE.Denominazione " + vbCrLF + _
	  "                         FROM VIEW_valid_testata_Strutture AGE WHERE AGE.RegCode = tb_strutture.Cod_Proprietario) " + vbCrLF + _
	  "               ELSE tb_strutture.denominazione END) as denominazione, " + vbCrLF + _
	  "         tipoimmobile, tipo " + vbCrLF + _
	  "     FROM tb_strutture INNER JOIN tb_loginStru ON tb_strutture.Str_ID = tb_loginStru.CURRENT_valid_STR_ID " + vbCrLF + _
	  "         INNER JOIN tb_modelli ON tb_loginStru.Modello=tb_modelli.Mod_ID " + vbCrLF + _
	  "     WHERE (AptCode=@APT OR AptCode LIKE '%' + @APT + '%') AND " + vbCrLF + _
	  "         (Mod_Tipo_record='S' OR Mod_Tipo_record='U' OR Mod_tipo_record='A') " + vbCrLF + _
	  "     ORDER BY datamodifica DESC " + vbCrLF + _
	  " ; " + _
	  " CREATE  PROCEDURE dbo.spWS_Hotel_Servizi " + vbCrLF + _
	  "     ( @REGCODE VARCHAR(12) ) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     DECLARE @STR_ID int		--id della struttura o dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @PRO_ID int		--id del proprietario dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @TYP_ID int		--id della tipologia dell'unit abitativa " + vbCrLF + _
	  "     DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLF + _
	  vbCrLf + _
	  "     SELECT @STR_ID = Current_valid_Str_id, " + vbCrLF + _
	  "         @TIPO_RECORD = mod_tipo_record " + vbCrLF + _
	  "     FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginStru.Modello=tb_modelli.mod_id " + vbCrLF + _
	  "     WHERE CodAlb = @REGCODE " + vbCrLF + _
	  vbCrLf + _
	  "     IF (@TIPO_RECORD='U') " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             DECLARE @PRO_REGCODE nvarchar(12) " + vbCrLF + _
	  "             DECLARE @TYP_REGCODE nvarchar(12) " + vbCrLF + _
	  "             SELECT @PRO_REGCODE=Cod_Proprietario, " + vbCrLF + _
	  "                 @TYP_REGCODE=Cod_Tipologia " + vbCrLF + _
	  "             FROM tb_Strutture " + vbCrLF + _
	  "             WHERE STR_ID = @STR_ID " + vbCrLF + _
	  vbCrLF + _
	  "             SELECT @PRO_ID = Current_valid_str_id " + vbCrLF + _
	  "             FROM tb_LoginStru " + vbCrLF + _
	  "             WHERE CodAlb = @PRO_REGCODE " + vbCrLF + _
	  vbCRLF + _
	  "             IF (@TYP_REGCODE IS NULL) " + vbCrLF + _
	  "                 SET @TYP_ID = 0 " + vbCrLF + _
	  "             ELSE " + vbCrLF + _
	  "                 SELECT @TYP_ID = Current_valid_str_id " + vbCrLF + _
	  "                     FROM tb_LoginStru " + vbCrLF + _
	  "                     WHERE CodAlb = @TYP_REGCODE " + vbCrLF + _
	  "         END " + vbCrLF + _
	  "     ELSE " + vbCrLF + _
	  "         BEGIN " + vbCrLF + _
	  "             SET @PRO_ID = 0 " + vbCrLF + _
	  "             SET @TYP_ID = 0 " + vbCrLF + _
	  "         END " + vbCrLF + _
	  vbCrLF + _
	  "     SELECT (tb_pubblicazioni_APT.pub_label_it) AS gruppo_it, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_en) AS gruppo_en, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_fr) AS gruppo_fr, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_de) AS gruppo_de, " + vbCrLF + _
	  "         (tb_pubblicazioni_APT.pub_label_es) AS gruppo_es, " + vbCrLF + _
	  "         (tb_servizi.serv_APT_nome_ITA) AS nome_it, " + vbCrLF + _
	  "         (tb_servizi.serv_APT_nome_ENG) AS nome_en, " + vbCrLF + _
	  "         (tb_servizi.serv_APT_nome_FRA) AS nome_fr, " + vbCrLF + _
	  "         (tb_servizi.serv_APT_nome_TED) AS nome_de, " + vbCrLF + _
	  "         (tb_servizi.serv_APT_nome_SPA) AS nome_es, " + vbCrLF + _
	  "         (tb_servizi.serv_symb) AS simbolo, " + vbCrLF + _
	  "         (rel_str_serv.rel_Str_serv_val) AS valore " + vbCrLF + _
	  "     FROM ((tb_servizi INNER JOIN rel_grp_serv ON tb_Servizi.serv_id=rel_grp_serv.rel_grp_id_serv) " + vbCrLF + _
	  "         INNER JOIN rel_str_serv ON rel_grp_serv.rel_grp_serv_id=rel_Str_serv.rel_str_id_relserv) " + vbCrLF + _
	  "         INNER JOIN tb_pubblicazioni_APT ON rel_grp_serv.serv_APT_pubblicazione=tb_pubblicazioni_APT.pub_id " + vbCrLF + _
	  "     WHERE (rel_str_serv.rel_id_str_serv = @STR_ID OR " + vbCrLF + _
	  "         rel_str_serv.rel_id_str_serv = @PRO_ID OR " + vbCrLF + _
	  "         rel_str_serv.rel_id_str_serv = @TYP_ID) " + vbCrLF + _
	  "         AND tb_pubblicazioni_APT.pub_order>0 " + vbCrLF + _
	  "         AND ((tb_servizi.serv_val=0) OR (rel_str_serv.rel_Str_serv_val<>'' AND rel_str_serv.rel_Str_serv_val IS NOT NULL)) " + vbCrLF + _
	  "     ORDER BY tb_pubblicazioni_APT.pub_order, serv_val, tb_servizi.serv_symb DESC " + vbCrLF + _
	  " ; " + _
	  " CREATE   PROCEDURE dbo.spWS_Hotel_Updated( " + vbCrLF + _
	  "     @APT nvarchar(2), " + vbCrLF + _
	  "     @DATAMOD nvarchar(10) " + vbCrLF + _
	  "     ) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     SELECT regcode, " + vbCrLF + _
	  "         (CASE WHEN mod_tipo_record='S' OR mod_tipo_record='A' THEN tb_strutture.Datamodifica " + vbCrLF + _
	  "               ELSE (SELECT MAX(tb_loginstru.current_valid_DataModifica) " + vbCrLF + _
	  "                         FROM tb_loginstru " + vbCrLF + _
	  "                         WHERE tb_loginstru.CodAlb = tb_strutture.RegCode OR " + vbCrLF + _
	  "                               tb_loginstru.codAlb = tb_strutture.Cod_Proprietario OR " + vbCrLF + _
	  "                               tb_loginstru.codalb = tb_strutture.Cod_Tipologia) END) AS datamodifica, " + vbCrLF + _
	  "         (CASE WHEN (mod_tipo_record='U' AND ISNULL(Denominazione,'')='') " + vbCrLF + _
	  "               THEN (SELECT AGE.Denominazione FROM VIEW_valid_testata_Strutture AGE " + vbCrLF + _
	  "                     WHERE AGE.RegCode=tb_strutture.Cod_Proprietario) " + vbCrLF + _
	  "               WHEN (modello=36) " + vbCrLF + _
	  "               THEN (SELECT tb_Strutture.Denominazione + ' - ' + AGE.Denominazione " + vbCrLF + _
	  "                     FROM VIEW_valid_testata_Strutture AGE " + vbCrLF + _
	  "                     WHERE AGE.RegCode=tb_strutture.Cod_Proprietario) " + vbCrLF + _
	  "               ELSE tb_strutture.denominazione END) as denominazione, " + vbCrLF + _
	  "         tipoimmobile, tipo " + vbCrLF + _
	  "     FROM tb_strutture INNER JOIN tb_loginStru ON tb_strutture.Str_ID = tb_loginStru.CURRENT_valid_STR_ID " + vbCrLF + _
	  "         INNER JOIN tb_modelli ON tb_loginStru.Modello=tb_modelli.Mod_ID " + vbCrLF + _
	  "     WHERE (AptCode=@APT OR AptCode LIKE '%' + @APT + '%') AND " + vbCrLF + _
	  "         (Mod_Tipo_record='S' OR Mod_Tipo_record='U' OR Mod_Tipo_Record='A') " + vbCrLF + _
	  "         AND (CASE WHEN mod_tipo_record='S' OR mod_tipo_record='A' THEN tb_strutture.Datamodifica " + vbCrLF + _
	  "                   ELSE (SELECT MAX(tb_loginstru.current_valid_DataModifica) " + vbCrLF + _
	  "                         FROM tb_loginstru " + vbCrLF + _
	  "                         WHERE tb_loginstru.CodAlb = tb_strutture.RegCode OR " + vbCrLF + _
	  "                               tb_loginstru.CodAlb = tb_strutture.Cod_Proprietario OR " + vbCrLF + _
	  "                               tb_loginstru.CodAlb = tb_strutture.Cod_Tipologia) END) >= CONVERT(DATETIME, @DATAMOD, 103) " + vbCrLF + _
	  "     ORDER BY datamodifica DESC " + vbCrLF + _
	  " ; " + _
	  " CREATE PROCEDURE dbo.spWS_Last_Update(@REGCODE nvarchar(12)) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     SELECT (SELECT MAX(tb_loginstru.current_valid_DataModifica) " + vbCrLF + _
	  "                 FROM tb_loginstru " + vbCrLF + _
	  "                 WHERE tb_loginstru.CodAlb = @REGCODE " + vbCrLF + _
	  "                     OR tb_loginstru.CodAlb = VIEW_valid_Testata_Strutture.Cod_Proprietario " + vbCrLF + _
	  "                     OR tb_loginstru.CodAlb = VIEW_valid_Testata_Strutture.Cod_Tipologia) AS datamodifica " + vbCrLF + _
	  "         FROM VIEW_valid_Testata_Strutture " + vbCrLF + _
	  "         WHERE RegCode=@REGCODE "
	  CALL DB.Execute(sql, 281)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 282
'...........................................................................................
'	Aggiunge campi descrizione ed intestazione modello per la compilazione
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "		mod_intestazione_premessa TEXT NULL, " + _
	  "		mod_intestazione_label_salva TEXT NULL, " + _
	  "		mod_intestazione_label_anteprima TEXT NULL, " + _
	  "		mod_intestazione_label_presenta TEXT NULL, " + _
	  "		mod_intestazione_label_chiudi TEXT NULL, " + _
	  "		mod_intestazione_istruzioni TEXT NULL " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_intestazione_premessa='', " + _
	  "		mod_intestazione_label_salva = 'Per salvare temporaneamente le modifiche e poter in seguito procedere ad ulteriori aggiornamenti online entro i termini di legge.', " + _
	  "		mod_intestazione_label_anteprima = 'Per generare una stampa di prova con gli ultimi dati salvati.', " + _
	  "		mod_intestazione_label_presenta = 'Se i dati sono completi e corretti &egrave;; possibile procedere alla presentazione della dichiarazione " + _
	  										  "che verr&agrave;; poi verificata e validata dagli operatori dell''amministrazione provinciale.<br>" + _
											  "<b>ATTENZIONE!!</b> La dichiarazione pu&ograve;; essere presentata una sola volta! <br>" + _
											  "Alla presentazione la procedura impedisce ulteriori modifiche ai dati, sar&agrave;; comunque possibile " + _
											  "stampare una anteprima della dichiarazione.', " + _
	  "		mod_intestazione_label_chiudi = 'Per uscire da questa sezione lasciando inalterati i dati precedentemente salvati.', " + _
	  "		mod_intestazione_istruzioni = 'I dati di cui &egrave;; possibile effettuare la modifica sono contenuti nelle aree il cui titolo &egrave;; evidenziato in rosso su sfondo grigio.<br>" + _
	  									  "Per i dati non modificabili per i quali fossero intervenute variazioni, vanno indicate nel campo ""NOTE"" posto alla fine della scheda.<br>" + _
										  "In caso di problemi rivolgersi seguente indirizzo <A href=""mailto:turismo@provincia.venezia.it"" class=""linkmodulo"">turismo@provincia.venezia.it</A>, " + _
										  "oppure telefonare ad uno dei seguenti numeri:'" + _
	  "		WHERE mod_dichiarazione_online=1 " + _
	  " ; "
CALL DB.Execute(sql, 282)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 283
'...........................................................................................
'	aggiunge campo per gestione notifiche via email delle dichiarazioni
'...........................................................................................
sql = " ALTER TABLE tb_turismo_admin_sito ADD " + _
	  "		tas_ricezione_notifiche BIT NULL " + _
	  " ; "
CALL DB.Execute(sql, 283)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 284
'...........................................................................................
'	statistiche - modifica campo sorgente
'...........................................................................................
sql = " ALTER TABLE rel_turismo_statCampiModelli ADD " + _
	  "		rcm_sorgente_new NVARCHAR(1000) NULL " + _
	  " ; "+ _
	  " UPDATE rel_turismo_statCampiModelli SET rcm_sorgente_new = CONVERT(NVARCHAR(1000), rcm_sorgente);"+ _
	  " ALTER TABLE rel_turismo_statCampiModelli DROP COLUMN rcm_sorgente;"+ _
	  " ALTER TABLE rel_turismo_statCampiModelli ADD"+ _
	  " 	rcm_sorgente NVARCHAR(1000) NULL;"+ _
	  " UPDATE rel_turismo_statCampiModelli SET rcm_sorgente = rcm_sorgente_new;"+ _
	  " ALTER TABLE rel_turismo_statCampiModelli DROP COLUMN rcm_sorgente_new;"
CALL DB.Execute(sql, 284)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 285
'...........................................................................................
'	converte dotazione "Pasto in italia - id=416" da flag+prezzo a flag+testo
'...........................................................................................
sql = " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=416;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=416; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=416) " + _
	  "			  AND rel_str_dotaz_pos_val>1 "
CALL DB.Execute(sql, 285)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 286
'...........................................................................................
'	converte dotazioni da flag+prezzo a flag+testo
'			 Pasto all'estero																dotaz_id = 415
'			 Pick-up in localit difficili da raggiungere 									dotaz_id = 417
'			 Pick-up fuori del centro storico			 									dotaz_id = 388
'			 Supplemento per lingua aggiuntiva 												dotaz_id = 387
'			 Supplemento per ogni persona in pi 											dotaz_id = 386
'			 Servizi forniti a clienti individuali o nuclei famigliari - tariffe 			dotaz_id = 457
'			 Servizi forniti a studenti italiani - tariffe 									dotaz_id = 459
'			Servizi forniti ad uno stesso committente in via continuativa - tariffe 		dotaz_id = 460
'...........................................................................................
sql = " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=415;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=415; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=415) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=417;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=417; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=417) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	   _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=388;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=388; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=388) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=387;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=387; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=387) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=386;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=386; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=386) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=457;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=457; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=457) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=459;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=459; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=459) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  _
	  " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=460;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=460; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=460) " + _
	  "			  AND rel_str_dotaz_pos_val>1 "
CALL DB.Execute(sql, 286)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 287
'...........................................................................................
'	converte dotazione:
'	Supplemento () per ogni ora o frazione oltre lorario max per ciascun servizio 
'	dotaz_id = 409
'	unifica e rimuove inoltre la dotazione "Supplemento (%) per ogni ora o frazione oltre lorario max per ciascun servizio "
'	con dotaz_id=474
'...........................................................................................
sql = " UPDATE tb_dotazioni SET " + _
	  "		dotaz_typ='" + TYPE_FLAG_TESTO + "', " + _
	  "		dotaz_nome_it = 'Supplemento per ogni ora o frazione oltre l''orario max per ciascun servizio' " + _
	  "		WHERE dotaz_id=409;" + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=409) " + _
	  "			  AND rel_str_dotaz_pos_val>1 " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' %', " + _
	  "		rel_str_dotaz_pos_val = rel_str_dotaz_pos_val + 1 " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=474) " + _
	  " UPDATE rel_grp_dotaz SET " + _
	  "		rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)', " + _
	  "		rel_grp_id_dotaz = 409 " + _
	  "		 WHERE rel_grp_id_dotaz=409 OR rel_grp_id_dotaz=474; " + _
	  "	DELETE FROM tb_dotazioni WHERE dotaz_id=474 "
CALL DB.Execute(sql, 287)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 288
'...........................................................................................
'	converte dotazione Altro rimborso 			dotaz_id = 429 in TYPE_FLAG_TESTO
'...........................................................................................
sql = " UPDATE tb_dotazioni SET dotaz_typ='" + TYPE_FLAG_TESTO + "' WHERE dotaz_id=429;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=429; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' &euro;;' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=429) " + _
	  "			  AND rel_str_dotaz_pos_val>2 "
CALL DB.Execute(sql, 288)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 289
'...........................................................................................
'	converte dotazioni a TYPE_TESTO
'		Supplemento per servizio notturno 													dotaz_id = 412
'		Supplemento per servizi resi durante le festivit 									dotaz_id = 414
'		Per servizi di particolare impegno o che richiedono una specifica preparazione		dotaz_id = 413
'...........................................................................................
sql = " UPDATE tb_dotazioni SET " + _
	  "			dotaz_typ='" + TYPE_TESTO + "', " + _
	  "			dotaz_lbl_1_1='min.', " + _
	  "			dotaz_lbl_1_2='max.', " + _
	  "			dotaz_lbl_1_3='unico' " + _
	  "		WHERE dotaz_id=412;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=412; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' %' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=412) " + _
	  _
	  " UPDATE tb_dotazioni SET " + _
	  "			dotaz_typ='" + TYPE_TESTO + "', " + _
	  "			dotaz_lbl_1_1='min.', " + _
	  "			dotaz_lbl_1_2='max.', " + _
	  "			dotaz_lbl_1_3='unico' " + _
	  "		WHERE dotaz_id=414;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=414; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' %' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=414) " + _
	  _
	  " UPDATE tb_dotazioni SET " + _
	  "			dotaz_typ='" + TYPE_TESTO + "', " + _
	  "			dotaz_lbl_1_1='min.', " + _
	  "			dotaz_lbl_1_2='max.', " + _
	  "			dotaz_lbl_1_3='unico' " + _
	  "		WHERE dotaz_id=413;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=413; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' %' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=413) "
CALL DB.Execute(sql, 289)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 290
'...........................................................................................
'	converte dotazione in TYPE_TESTO_LINEA
'		Condizione e percentuale di aumento/riduzione tariffa  								dotaz_id = 391
'...........................................................................................
sql = " UPDATE tb_dotazioni SET " + _
	  "			dotaz_typ='" + TYPE_TESTO_LINEA + "', " + _
	  "			dotaz_lbl_1_2 = 'min.', " + _
	  "			dotaz_lbl_2_1 = 'max', " + _
	  "			dotaz_lbl_2_2 = 'unico', " + _
	  "			dotaz_lbl_1_3 = '', " + _
	  "			dotaz_lbl_1_4 = '' " + _
	  "		WHERE dotaz_id=391;" + _
	  " UPDATE rel_grp_dotaz SET rel_grp_dotaz_note = 'Indicare l''unit&agrave;; di misura di riferimento (&euro;; / %)' WHERE rel_grp_id_dotaz=391; " + _
	  " UPDATE rel_str_dotaz SET " + _
	  "		rel_str_dotaz_testo_it = CAST(rel_str_dotaz_valore AS nvarchar(50)) + ' %' " + _
	  "		WHERE rel_str_id_dotaz IN( SELECT rel_Grp_dotaz_id FROM rel_grp_dotaz WHERE rel_grp_id_dotaz=391) " + _
	  "			  AND rel_str_dotaz_pos_val>1 "
CALL DB.Execute(sql, 290)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 291
'...........................................................................................
'	aggiorna colonna nome dotazione
'...........................................................................................
sql = " ALTER TABLE tb_dotazioni ALTER COLUMN dotaz_nome_it nvarchar(250) NULL "
CALL DB.Execute(sql, 291)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 292
'...........................................................................................
'	configurazioni di default
'		aggiunge campo file per firma dirigente dei tesserini
'		aggiunge campo file per vidimazione su modelli stampati
'...........................................................................................
sql = " ALTER TABLE tb_modelli_default ADD " + _
	  "		default_immagine_vidimazione_tesserini nvarchar(250) NULL, " + _
	  "		default_immagine_vidimazione_modelli nvarchar(250) NULL " + _
	  " ; " + _
	  " UPDATE tb_modelli_default SET default_immagine_vidimazione_tesserini = 'tesserini_firma_dirigente.gif' "
CALL DB.Execute(sql, 292)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 293
'...........................................................................................
'	aggiunge campi per gestione email
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "		mod_email_sender nvarchar(250), " + _
	  "		mod_email_validazione_oggetto nvarchar(250), " + _
	  "		mod_email_conferma_oggetto nvarchar(250), " + _
	  "		mod_email_password_oggetto nvarchar(250) " + _
	  " ; " + _
	  " UPDATE tb_modelli SET " + _
	  "		mod_email_sender='turismo@provincia.venezia.it', " + _
	  "		mod_email_conferma_oggetto = 'Dichiarazione presentata - www.turismo.provincia.venezia.it', " + _
	  "		mod_email_validazione_oggetto = 'Dichiarazione completata e validata - www.turismo.provincia.venezia.it', " + _
	  "		mod_email_password_oggetto = 'Accesso area riservata - www.turismo.provincia.venezia.it' "
CALL DB.Execute(sql, 293)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 294
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__6(conn)
CALL DB.Execute(sql, 294)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 295
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__7(conn)
CALL DB.Execute(sql, 295)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 296
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__8(conn)
CALL DB.Execute(sql, 296)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 297
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__9(conn)
CALL DB.Execute(sql, 297)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 298
'...........................................................................................
'	aggiunge campi per gestione storico
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		archivio_modello_dichiarazione nvarchar(250) NULL, " + _
	  "		archivio_tabella_prezzi nvarchar(250) NULL " + _
	  " ; "
CALL DB.Execute(sql, 298)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 299
'...........................................................................................
'	genera modelli storici per professioni turistiche
'...........................................................................................
sql = " SELECT * FROM AA_Versione "
CALL DB.Execute(sql, 299)
if DB.last_update_executed then
	CALL Aggiornamento_299_GenerazioneStoricoModelliPT(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_299_GenerazioneStoricoModelliPT(conn, rs)
	dim sql, path, x_sql
	dim fso, tempEmail, tempStream
	dim FileModelloDichiarazione, FileTabellaPrezzi
	
	set tempEmail = new mailer
	x_sql = ""
	sql = "SELECT *, " + _
		  " (SELECT mod_directory FROM tb_modelli INNER JOIN tb_loginstru ON tb_modelli.mod_id = tb_loginstru.modello " + _
		  "			WHERE tb_loginstru.CodAlb = tb_strutture.RegCode ) AS directory_modello " + _
		  " FROM tb_strutture WHERE " + _
		  " IsNull(anno_prezzi, 0)>=2005 AND " + _
		  " IsNull(archivio_modello_dichiarazione, '')='' AND " + _
		  " RegCode IN (SELECT CodAlb FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginstru.modello=tb_modelli.mod_id " + _
		  "				WHERE mod_tipo_record='" & RECORD_TYPE_PT & "' AND mod_dichiarazione_tipo<>'" + DICHIARAZIONE_NESSUNA + "' ) " + _
		  " AND DataModifica < " + SQL_date(conn, DateSerial(2006, 06, 29)) + _
		  " ORDER BY RegCode, str_id "
	response.write "<!-- " + vbCrLf + sql + vbCrLf + " -->" + vbCrLf
	rs.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
	while not rs.eof
	
		response.write "<!-- " + vbCrLf + _
					   "regcode=" + rs("RegCode") + vbCrLf + _
					   "str_id=" & rs("str_id") & vbCrLf + _
					   vbCrLf + " -->" + vbCrLf
		
		'carica html modello
		tempEmail.LoadHTML "http://" & Application("SERVER_NAME") & "/riservata/" & rs("directory_modello") & _
						   "/Modulo_stampa_storico.asp?str_id=" & rs("str_id") & "&ARCHIVIAZIONE=1", _
						   "http://" & Application("SERVER_NAME") & "/riservata"
		
		'recupera contenuto modello
		set tempStream = tempEmail.Message.HTMLBodyPart.GetDecodedContentStream
		
		'salva modello su file
		FileModelloDichiarazione = rs("regcode") & "_" & rs("str_id") & "_modello_dichiarazione.htm"
		path = Application("IMAGE_PATH") & "dichiarazioni_strutture\" + FileModelloDichiarazione
		tempStream.SaveToFile path, 2
		
		'imopsta i dati del modello validato
		rs("archivio_modello_dichiarazione") = FileModelloDichiarazione
		rs("archivio_tabella_prezzi") = ""
		rs.update
		
		rs.movenext
	wend
	rs.close
	
	if x_sql<>"" then
		CALL conn.execute(x_sql)
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 300
'...........................................................................................
'	corregge dati presenti in "AptCode" delle strutture
'...........................................................................................
sql = " UPDATE tb_strutture SET AptCode = LTRIM(RTRIM(AptCode)) WHERE AptCode IS NOT NULL "
CALL DB.Execute(sql, 300)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 301
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__10(conn)
CALL DB.Execute(sql, 301)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 302
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__11(conn)
CALL DB.Execute(sql, 302)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 303
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__12(DB, rs, 303)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 304
'...........................................................................................
'	crea funzione per gestione dati nello storico
'...........................................................................................
sql =	" CREATE FUNCTION dbo.fn_valid_strutture ( @anno int ) " + vbCrLF + _
		"   RETURNS table " + vbCrLF + _
		"   AS RETURN ( " + vbcRLF + _
		"   SELECT tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND " + vbCrLf + _
		"                           str_id IN (SELECT TOP 1 str_id " + vbCrLf + _
		"                                      FROM tb_strutture tb_storico " + vbCrLF + _
		"                                      WHERE tb_storico.RegCode = tb_strutture.RegCode " + vbCrLF + _
		"                                            AND IsNull(tb_storico.anno_prezzi,Year(tb_storico.DataModifica))<=@anno " + vbCrLF + _
		"                                            AND ( str_id IN (SELECT str_id FROM tb_strutture tb_sub_storico WHERE tb_sub_storico.anno_prezzi = @anno AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                  OR " + vbCrLF + _
		"                                                  ( NOT EXISTS(SELECT str_id FROM tb_strutture tb_sub_storico WHERE tb_sub_storico.anno_prezzi = @anno AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                    AND " + vbCrLF + _
		"                                                    str_id IN (SELECT str_id FROM tb_strutture tb_sub_storico WHERE tb_sub_storico.DataModifica < CONVERT(DATETIME, CAST(@anno AS nvarchar(4)) + '-12-31 23:59:59', 102) AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                  ) " + vbCrLF + _
		"                                                ) " + vbCrlf + _
		"                                      ORDER BY tb_storico.regcode, tb_storico.DataModifica DESC, tb_storico.str_id DESC " + vbCrLF + _
		"                                     ) INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
		"   ) " + vbCrLF + _
        " ; "
CALL DB.Execute(sql, 304)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 305
'...........................................................................................
'aggiorna stato proprietari ed agenzie
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 305)
if DB.last_update_executed then
	CALL Aggiornamento_305_AggiornaStatoAttivitaProprietari(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_305_AggiornaStatoAttivitaProprietari(DbConn, rs)
	dim readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.Open Application(request("ConnString")), "", ""
	
	sql = " SELECT CodAlb FROM tb_loginstru INNER JOIN tb_modelli ON tb_loginstru.modello = tb_modelli.mod_id " + _
		  " WHERE mod_tipo_record IN ('" & RECORD_TYPE_AGENCY & "', '" & RECORD_TYPE_OWNER & "') "
	readRs.open sql, readConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	while not readRs.eof
		CALL Conn.spstr_ATTIVITA_PROPRIETARIO(readRs("CodAlb"))
		readRs.movenext
	wend	
	readRs.close
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 306
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__13(conn)
CALL DB.Execute(sql, 306)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 307
'...........................................................................................
sql = rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
CALL DB.Execute(sql, 307)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 308
'...........................................................................................
sql = Install__FRAMEWORK_CORE__NEXTWEB5(conn)
CALL DB.Execute(sql, 308)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 309
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__14(conn)
CALL DB.Execute(sql, 309)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 310
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__15(conn)
CALL DB.Execute(sql, 310)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 311
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__16(conn)
CALL DB.Execute(sql, 311)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 312
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__17(conn)
CALL DB.Execute(sql, 312)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 313
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__18(conn)
CALL DB.Execute(sql, 313)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 314
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__19(conn)
CALL DB.Execute(sql, 314)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 315
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__20(conn)
CALL DB.Execute(sql, 315)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 316
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__21(conn)
CALL DB.Execute(sql, 316)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 317
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__22(conn)
CALL DB.Execute(sql, 317)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 318
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__23(conn)
CALL DB.Execute(sql, 318)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 319
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__24(conn)
CALL DB.Execute(sql, 319)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 320
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__25(conn)
CALL DB.Execute(sql, 320)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 321
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__26(conn) + _
	  Aggiornamento__FRAMEWORK_CORE__27(conn)
CALL DB.Execute(sql, 321)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 322
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__28(conn)
CALL DB.Execute(sql, 322)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 323
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__29(conn)
CALL DB.Execute(sql, 323)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 324
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__30(conn)
CALL DB.Execute(sql, 324)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 325
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__31(conn)
CALL DB.Execute(sql, 325)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 326
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__32(conn)
CALL DB.Execute(sql, 326)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 327
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__33(conn)
CALL DB.Execute(sql, 327)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 328
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__34(conn)
CALL DB.Execute(sql, 328)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 329
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__35(conn)
CALL DB.Execute(sql, 329)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 330
'...........................................................................................
'	aggiornamento procedure per pubblicazione prezzi via webservice
'...........................................................................................
sql = DropObject(conn, "spWS_Hotel_Dotazioni", "PROCEDURE") + _
	  " CREATE  PROCEDURE dbo.spWS_Hotel_Dotazioni " + vbCrLF + _
	  "     ( @REGCODE VARCHAR(13)) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "     DECLARE @STR_ID int		--id della struttura o dell'unit abitativa " + vbCrLf + _
      "     DECLARE @PRO_ID int		--id del proprietario dell'unit abitativa " + vbCrLf + _
      "     DECLARE @TYP_ID int		--id della tipologia dell'unit abitativa " + vbCrLf + _
      "     DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLf + _
      "     DECLARE @ANNO_PREZZI int " + vbCrLf + _
      "     DECLARE @MODELLO int " + vbCrLf + _
      vbCrLf + _
      "     SELECT @STR_ID = current_valid_str_id, @TIPO_RECORD = mod_tipo_record, @MODELLO = modello, @ANNO_PREZZI = anno_prezzi " + vbCrLf + _
      "         FROM View_valid_strutture " + vbCrLf + _
      "         WHERE CodAlb = @REGCODE " + vbCrLf + _
      vbCrLF + _
      "     IF (@TIPO_RECORD='U') " + vbCrLf + _
      "         BEGIN " + vbCrLf + _
      "             DECLARE @PRO_REGCODE nvarchar(12) " + vbCrLf + _
      "             DECLARE @TYP_REGCODE nvarchar(12) " + vbCrLf + _
      vbCrLf + _
      "             SELECT @PRO_REGCODE=Cod_Proprietario, @TYP_REGCODE=Cod_Tipologia " + vbCrLf + _
      "                 FROM tb_Strutture " + vbCrLf + _
      "                 WHERE STR_ID=@STR_ID " + vbCrLf + _
      vbCrLF + _
      "             SELECT @PRO_ID = Current_valid_str_id , @ANNO_PREZZI = anno_prezzi " + vbCrLf + _
      "                 FROM view_valid_strutture " + vbCrLf + _
      "                 WHERE CodAlb = @PRO_REGCODE " + vbCrLf + _
      vbCrLF + _
      "             IF (@TYP_REGCODE IS NULL) " + vbCrLf + _
      "                 SET @TYP_ID = 0 " + vbCrLf + _
      "             ELSE " + vbCrLf + _
      "                 SELECT @TYP_ID = Current_valid_str_id " + vbCrLf + _
      "                     FROM tb_LoginStru " + vbCrLf + _
      "                     WHERE CodAlb = @TYP_REGCODE " + vbCrLf + _
      "             END " + vbCrLf + _
      "     ELSE " + vbCrLf + _
      "         BEGIN " + vbCrLf + _
      "             SET @PRO_ID = 0 " + vbCrLf + _
      "             SET @TYP_ID = 0 " + vbCrLf + _
      "         END " + vbCrLf + _
      vbCRlf + _
      "     SELECT (tb_pubblicazioni_APT.pub_label_it) AS gruppo_it, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_en) AS gruppo_en, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_fr) AS gruppo_fr, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_de) AS gruppo_de, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_es) AS gruppo_es, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_APT_nome_ITA) AS nome_it, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_APT_nome_ENG) AS nome_en, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_APT_nome_FRA) AS nome_fr, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_APT_nome_TED) AS nome_de, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_APT_nome_SPA) AS nome_es, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_symb) AS simbolo, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_typ) AS tipo, " + vbCrLf + _
      "         (tb_dotazioni.dotaz_num_val) AS numero_valori, " + vbCrLf + _
      "         (dotaz_lbl_1_1) AS label_1_level_1, " + vbCrLf + _
      "         (dotaz_lbl_1_2) AS label_2_level_1, " + vbCrLf + _
      "         (dotaz_lbl_1_3) AS label_3_level_1, " + vbCrLf + _
      "         (dotaz_lbl_1_4) AS label_4_level_1, " + vbCrLf + _
      "         (dotaz_lbl_2_1) AS label_1_level_2, " + vbCrLf + _
      "         (dotaz_lbl_2_2) AS label_2_level_2, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' OR dotaz_typ='MN' OR dotaz_typ='MP' THEN rel_1.rel_str_dotaz_testo_it ELSE CAST(rel_1.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore1, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' OR dotaz_typ='FM' THEN rel_2.rel_str_dotaz_testo_it ELSE CAST(rel_2.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore2, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_3.rel_str_dotaz_testo_it ELSE CAST(rel_3.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore3, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_4.rel_str_dotaz_testo_it ELSE CAST(rel_4.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore4, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_5.rel_str_dotaz_testo_it ELSE CAST(rel_5.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore5, " + vbCrLf + _
      "         (ISNULL(CASE WHEN dotaz_typ='T' THEN rel_6.rel_str_dotaz_testo_it ELSE CAST(rel_6.rel_str_dotaz_valore AS nvarchar(50)) END, '')) AS valore6," + vbCrLf + _
      "         pub_order, dotaz_apt_ordine " + vbCrLf + _
      "         FROM tb_dotazioni " + vbCrLf + _
      "             INNER JOIN rel_grp_dotaz ON tb_dotazioni.dotaz_id=rel_grp_dotaz.rel_grp_id_dotaz " + vbCrLf + _
      "             INNER JOIN tb_pubblicazioni_APT ON dotaz_APT_pubblicazione = tb_pubblicazioni_APT.pub_id " + vbCrLf + _
      "             INNER JOIN  tb_grp_vis ON rel_grp_dotaz.rel_id_grp_dotaz = tb_grp_vis.grp_id" + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_1 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_1.rel_str_id_dotaz AND rel_1.rel_str_dotaz_pos_val=1 AND (rel_1.rel_id_str_dotaz=@STR_ID OR rel_1.rel_id_str_dotaz=@PRO_ID OR rel_1.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_2 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_2.rel_str_id_dotaz AND rel_2.rel_str_dotaz_pos_val=2 AND (rel_2.rel_id_str_dotaz=@STR_ID OR rel_2.rel_id_str_dotaz=@PRO_ID OR rel_2.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_3 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_3.rel_str_id_dotaz AND rel_3.rel_str_dotaz_pos_val=3 AND (rel_3.rel_id_str_dotaz=@STR_ID OR rel_3.rel_id_str_dotaz=@PRO_ID OR rel_3.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_4 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_4.rel_str_id_dotaz AND rel_4.rel_str_dotaz_pos_val=4 AND (rel_4.rel_id_str_dotaz=@STR_ID OR rel_4.rel_id_str_dotaz=@PRO_ID OR rel_4.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_5 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_5.rel_str_id_dotaz AND rel_5.rel_str_dotaz_pos_val=5 AND (rel_5.rel_id_str_dotaz=@STR_ID OR rel_5.rel_id_str_dotaz=@PRO_ID OR rel_5.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "             LEFT JOIN rel_str_dotaz rel_6 ON rel_grp_dotaz.rel_grp_dotaz_id = rel_6.rel_str_id_dotaz AND rel_6.rel_str_dotaz_pos_val=6 AND (rel_6.rel_id_str_dotaz=@STR_ID OR rel_6.rel_id_str_dotaz=@PRO_ID OR rel_6.rel_id_str_dotaz=@TYP_ID) " + vbCrLf + _
      "         WHERE tb_pubblicazioni_APT.pub_order>0 " + vbCrLf + _
      "               AND rel_grp_dotaz_id IN (SELECT rel_str_id_dotaz FROM rel_str_dotaz WHERE (rel_id_str_dotaz=@STR_ID OR rel_id_str_dotaz=@PRO_ID OR rel_id_str_dotaz=@TYP_ID) AND ((rel_str_dotaz_valore <> 0 AND rel_Str_dotaz_valore IS NOT NULL) OR (rel_str_dotaz_testo_it <> '' AND rel_Str_dotaz_testo_it IS NOT NULL))) " + vbCrLf + _
      "               AND ( Grp_Admin_page<>'P' OR @MODELLO<>31 OR @ANNO_PREZZI = Year(GETDATE()) ) " + vbCrLf + _
      "     UNION " + vbCrLf + _
      "     SELECT (tb_pubblicazioni_APT.pub_label_it) AS gruppo_it, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_en) AS gruppo_en, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_fr) AS gruppo_fr, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_de) AS gruppo_de, " + vbCrLf + _
      "         (tb_pubblicazioni_APT.pub_label_es) AS gruppo_es, " + vbCrLf + _
      "         ('Prezzi validi per l''anno ' + CAST(@ANNO_PREZZI AS nvarchar(10))) AS nome_it, " + vbCrLf + _
      "         ('Prices for the year ' + CAST(@ANNO_PREZZI AS nvarchar(10))) AS nome_en, " + vbCrLf + _
      "         ('') AS nome_fr, " + vbCrLf + _
      "         ('') AS nome_de, " + vbCrLf + _
      "         ('') AS nome_es, " + vbCrLf + _
      "         ('') AS simbolo, " + vbCrLf + _
      "         ('P') AS tipo, " + vbCrLf + _
      "         (1) AS numero_valori, " + vbCrLf + _
      "         ('') AS label_1_level_1, " + vbCrLf + _
      "         ('') AS label_2_level_1, " + vbCrLf + _
      "         ('') AS label_3_level_1, " + vbCrLf + _
      "         ('') AS label_4_level_1, " + vbCrLf + _
      "         ('') AS label_1_level_2, " + vbCrLf + _
      "         ('') AS label_2_level_2, " + vbCrLf + _
      "         (NULL) AS valore1, " + vbCrLf + _
      "         (NULL) AS valore2, " + vbCrLf + _
      "         (NULL) AS valore3, " + vbCrLf + _
      "         (NULL) AS valore4, " + vbCrLf + _
      "         (NULL) AS valore5, " + vbCrLf + _
      "         (NULL) AS valore6, " + vbCrLf + _
      "         pub_order, " + vbCrLf + _
      "         (100) AS dotaz_apt_ordine " + vbCrLf + _
      "     FROM tb_pubblicazioni_APT " + vbCrLf + _
      "     WHERE tb_pubblicazioni_APT.pub_label_it LIKE '%Prezzi%' " + vbCrLf + _
      "         AND IsNull(@ANNO_PREZZI,0)>0 " + vbCrLf + _
      "     ORDER BY pub_order, dotaz_APT_ordine, nome_it "
CALL DB.Execute(sql, 330)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 331
'...........................................................................................
'	corregge proprietari degli oggetti creati senza prefisso "dbo."
'...........................................................................................
sql = " sp_changeobjectowner 'turismo.v_indice','dbo' ; " + _
      " sp_changeobjectowner 'turismo.v_indice_visibile','dbo' ; " + _
      " sp_changeobjectowner 'turismo.tb_menuItem','dbo' ; " + _
      " sp_changeobjectowner 'turismo.tb_siti_tabelle_pubblicazioni','dbo' ; " + _
      " sp_changeobjectowner 'turismo.fn_records_strutture','dbo' ; " + _
      " sp_changeobjectowner 'turismo.fn_valid_strutture','dbo' ; "
CALL DB.Execute(sql, 331)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 332
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__36(conn)
CALL DB.Execute(sql, 332)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 333
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__37(conn)
CALL DB.Execute(sql, 333)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 334
'...........................................................................................
sql = AggiornamentoSpeciale__FRAMEWORK_CORE__38(DB, rs, 334)
CALL DB.Execute(sql, 334)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 335
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__39(conn)
CALL DB.Execute(sql, 335)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(335)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 336
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__40(conn)
CALL DB.Execute(sql, 336)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 337
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__41(conn)
CALL DB.Execute(sql, 337)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 338
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__42(conn)
CALL DB.Execute(sql, 338)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 339
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__43(conn)
CALL DB.Execute(sql, 339)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 340
'...........................................................................................
'	aggiunge colonne alle strutture dati dei modelli per lo scambio dati con rvtweb
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD mod_rvtweb_codice nvarchar(2); " + _
      " ALTER TABLE tb_tipi_str ADD tip_rvtweb_codice nvarchar(2); " + _
      " ALTER TABLE tb_localita ADD loc_rvtweb_codice nvarchar(3); "
CALL DB.Execute(sql, 340)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 341
'...........................................................................................
'	aggiunge colonna allo storico per la gestione del "progressivo"
'   corregge codice regionale
'...........................................................................................
sql = " ALTER TABLE tb_str_logs ADD str_log_progressivo INT NULL; " + _
      " UPDATE tb_str_logs SET str_log_codAlb = RTRIM(LTRIM(str_log_codAlb)); "
CALL DB.Execute(sql, 341)
if DB.last_update_executed then
	CALL Aggiornamento_341_GenerazioneProgressivo(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_341_GenerazioneProgressivo(conn, rs)
    dim CurrentRegCode, Progressivo
    sql = " SELECT * FROM tb_str_logs " + _
          " WHERE ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione%')) " + _
          " ORDER BY str_log_codAlb, str_log_id "
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    CurrentRegCode = ""
    Progressivo = 0
    while not rs.eof
        if CurrentRegCode <> Trim(rs("str_log_codAlb")) then
            CurrentRegCode = Trim(rs("str_log_codAlb"))
            Progressivo = 0
        else
            Progressivo = Progressivo + 1
        end if
        rs("str_log_progressivo") = Progressivo
        rs.Update
        
        rs.movenext
    wend
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 342
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__44(conn)
CALL DB.Execute(sql, 342)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 343
'...........................................................................................
'	aggiunge trigger alla struttura di log per il calcolo del progressivo
'...........................................................................................
sql = " CREATE TRIGGER dbo.tb_str_logs_INSERT ON tb_str_logs AFTER INSERT AS " + vbCrLf + _
      "     DECLARE @REGCODE nvarchar(13) " + vbCrLF + _
      "     DECLARE @LAST INT " + vbCrLF + _
      "     DECLARE @PROGRESSIVO INT " + vbCrLF + _
      vbCrLf + _
      "     SELECT @REGCODE = RTRIM(LTRIM(str_log_CodAlb)), @LAST=str_log_id " + vbCrLF + _
      "         FROM INSERTED WHERE ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione%')) " + vbCrLf + _
      vbCrLF + _
      "     if (@REGCODE <> '') " + vbCRLF + _
      "         BEGIN " + vbCrLF + _
      "             IF (EXISTS(SELECT * FROM tb_str_logs WHERE RTRIM(LTRIM(str_log_CodAlb)) LIKE @REGCODE AND str_log_id <> @LAST " + vbCrLF + _
      "                                                        AND ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione%')) )) " + vbCrLf + _
      "                 BEGIN " + vbCrLF + _
      "                     SELECT @PROGRESSIVO = MAX(str_log_progressivo) FROM tb_str_logs " + vbCrLF + _
      "                         WHERE RTRIM(LTRIM(str_log_CodAlb)) LIKE @REGCODE AND str_log_id <> @LAST " + vbCrLF + _
      "                               AND ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione%')) " + vbCrLf + _
      "                     SET @PROGRESSIVO = @PROGRESSIVO + 1 " + vbCrLF + _
      "                 END " + vbCrLf + _
      "             ELSE " + vbCrLF + _
      "                 BEGIN " + vbCrLF + _
      "                     SET @PROGRESSIVO = 0 " + vbCrLF + _
      "                 END " + vbCrLF + _
      "             UPDATE tb_str_logs SET str_log_progressivo=@PROGRESSIVO WHERE str_log_id=@LAST " + vbCrLF + _
      "         END " + vbCrLF + _
      " ; "
CALL DB.Execute(sql, 343)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 344
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__45(conn)
CALL DB.Execute(sql, 344)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 345
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__46(conn)
CALL DB.Execute(sql, 345)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 346
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__47(conn)
CALL DB.Execute(sql, 346)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 347
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__48(conn)
CALL DB.Execute(sql, 347)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 348
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__49(conn)
CALL DB.Execute(sql, 348)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 349
'...........................................................................................
sql = "SELECT * FROM aa_versione"
CALL DB.Execute(sql, 349)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 350
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__50(conn)
CALL DB.Execute(sql, 350)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 351
'...........................................................................................
'	aggiunge funzione per il calcolo del codice regionale per la regione.
'...........................................................................................
sql = "  CREATE FUNCTION dbo.fn_regcode_for_regione (@CODE nvarchar(13), @MODELLO int)  " + vbCrLF + _
      "     RETURNS nvarchar(13) " + vbCrLF + _
      " AS " + vbCrLF + _
      " BEGIN " + vbCrLF + _
      "     DECLARE @CODEREGIONE nvarchar(13) " + vbCrLF + _
      "     SELECT @CODEREGIONE = LTRIM(REPLACE( @CODE,  " + vbCrLF + _
      "                                          (SELECT rel_radice_regcode FROM rel_mod_apt_uffici WHERE rel_mod_id=@MODELLO AND @CODE LIKE rel_radice_regcode + '%' ), " + vbCrLF + _
      "                                          (SELECT rel_radice_regione FROM rel_mod_apt_uffici WHERE rel_mod_id=@MODELLO AND @CODE LIKE rel_radice_regcode + '%' ))) " + vbCrLF + _
      vbCrLF + _
      "     RETURN ISNULL(@CODEREGIONE, @CODE) " + vbCrLF + _
      " END "
CALL DB.Execute(sql, 351)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 352
'...........................................................................................
'	aggiunge funzione per il recupero del codice regionale da quello per la regione
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_regcode_from_regione (@CODEREGIONE nvarchar(13), @MODELLO int) " + vbCrLF + _
      "     RETURNS nvarchar(13) " + vbCrLF + _
      " AS " + vbCrLF + _
      " BEGIN " + vbCrLF + _
      "     DECLARE @CODE nvarchar(13) " + vbCrLF + _
      "     SELECT @CODE = LTRIM(REPLACE( @CODEREGIONE,  " + vbCrLF + _
      "                                   (SELECT rel_radice_regione FROM rel_mod_apt_uffici WHERE @MODELLO=rel_mod_id AND @CODEREGIONE LIKE rel_radice_regione + '%' ), " + vbCrLF + _
      "                                   (SELECT rel_radice_regcode FROM rel_mod_apt_uffici WHERE @MODELLO=rel_mod_id AND @CODEREGIONE LIKE rel_radice_regione + '%' ))) " + vbCrLF + _
      vbCrLF + _
      " 	RETURN ISNULL(@CODE, @CODEREGIONE) " + vbCrLF + _
      " END "
CALL DB.Execute(sql, 352)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 353
'...........................................................................................
'	aggiunge funzione per il calcolo del totale camere
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_calcola_totale_camere (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      " BEGIN " + vbCrLF + _
      "     /* " + vbCrLF + _
      "     Camere singole:			168, 342 " + vbCrLF + _
      "     Camere doppie:			169, 343 " + vbCrLF + _
      "     Camere a pi letti:		170, 344  " + vbCrLF + _
      "     Suite				172 " + vbCrLF + _
      "     Juniorsuite:			171 " + vbCrLF + _
      "     */  " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMERE int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMERE = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "         AND rel_grp_id_dotaz IN (168, 342, 169, 343, 170, 344, 172, 171) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Monolocali:		187, 220 " + vbCrLF + _
      "     Bilocali:		188, 221 " + vbCrLF + _
      "     Trilocali:		189, 222 " + vbCrLF + _
      "     Pi locali:		337, 223 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_UA int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (187, 220, 188, 221) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + SUM(rel_str_dotaz_valore * 2)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (189, 222) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + SUM(rel_str_dotaz_valore - 1)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 3 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (337, 223) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Totale Piazzole:	301 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMPEGGI int  " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (301) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + SUM((num_vani - 1) * qta_ua)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)>1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + SUM(num_vani * qta_ua)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)=1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      vbCrLF + _
      "     RETURN ISNULL(@TOT_CAMERE_CAMERE, 0) + ISNULL(@TOT_CAMERE_UA, 0) + ISNULL(@TOT_CAMERE_CAMPEGGI, 0) " + vbCrLF + _
      " END "
CALL DB.Execute(sql, 353)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 354
'...........................................................................................
'	aggiunge funzione per il calcolo del totale posti letto.
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_calcola_totale_posti_letto (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         /* " + vbCrLF + _
      "         Totale posti letto camere:		174 " + vbCrLF + _
      "         Totale posti letto suite:		176 " + vbCrLF + _
      "         Totale posti lestto u.a.:		195 " + vbCrLF + _
      "         CRM campeggi:				    308 " + vbCrLF + _
      "         */ " + vbCrLF + _
      vbCrLF + _
      "         DECLARE @TOT_PL int " + vbCrLF + _
      "         SELECT @TOT_PL = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND  rel_grp_id_dotaz IN (174, 176, 195, 308) " + vbCrLF + _
      "         RETURN @TOT_PL " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 354)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 355
'...........................................................................................
'	aggiunge funzione per il calcolo del totale bagni
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_calcola_totale_bagni (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      "  " + vbCrLF + _
      " AS  " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         /* " + vbCrLF + _
      "         Totale generale bagni:				207 " + vbCrLF + _
      "         Camerini bagno chiusi:				317 " + vbCrLF + _
      "         Servizi igienici per singoli equipaggi:		326 " + vbCrLF + _
      "         Servizi igienici per disabili:			318 " + vbCrLF + _
      "         Unit abitative con servizi igienici:		303 " + vbCrLF + _
      "         */ " + vbCrLF + _
      vbCrLF + _
      "         DECLARE @TOT_BAGNI int " + vbCrLF + _
      "         SELECT @TOT_BAGNI = SUM(rel_str_dotaz_valore)" + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND  rel_grp_id_dotaz IN (207, 317, 326, 318) " + vbCrLF + _
      vbCrLF + _
      "         SELECT @TOT_BAGNI = @TOT_BAGNI + SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val IN (1, 3)  AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND rel_grp_id_dotaz IN (303)  " + vbCrLF + _
      "         RETURN @TOT_BAGNI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 355)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 356
'...........................................................................................
'	aggiunge funzione per il calcolo degli esercizi
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_calcola_totale_esercizi (@STR_ID int, @MODELLO int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         DECLARE @TOT_ESERCIZI int " + vbCrLF + _
      vbCrLF + _
      "         if (@MODELLO = 21 OR @MODELLO=31 OR @MODELLO=36) " + vbCrLF + _
      "         BEGIN " + vbCrLF + _
      "             SELECT @TOT_ESERCIZI = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "                 FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "                 WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                       AND rel_grp_id_dotaz IN (194) " + vbCrLF + _
      "         END " + vbCrLF + _
      "         ELSE BEGIN " + vbCrLF + _
      "             SET @TOT_ESERCIZI = 1 " + vbCrLF + _
      "         END " + vbCrLF + _
      vbCrLF + _
      "         RETURN @TOT_ESERCIZI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 356)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 357
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__51(conn)
CALL DB.Execute(sql, 357)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 358
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__52(conn)
CALL DB.Execute(sql, 358)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 359
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__53(conn)
CALL DB.Execute(sql, 359)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 360
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__54(conn)
CALL DB.Execute(sql, 360)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 361
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__55(conn)
CALL DB.Execute(sql, 361)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 362
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__56(conn)
CALL DB.Execute(sql, 362)
'*******************************************************************************************


'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'NUOVI AGGIORNAMENTI PER INTEGRAZIONE PORTALI APT
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 363
'...........................................................................................
'aggiorna versione del portale alla nuova struttura dati
'...........................................................................................
sql = "UPDATE AA_versione SET Versione=994 "
CALL DB.Execute(sql, 363)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 996
'...........................................................................................
'creazione ed inizializzazione struttura per NEXT-INFO
'aggiunge campi per gestione bussola (SOLO PER NEXT-INFO APT)
'aggiunge campo per gestione validita' prezzi (descrittori anagrafiche)
'...........................................................................................
sql = Install__NEXTINFO(conn) + "; " + _
      Activate_NEXTINFO(conn) + "; " + _
		"CREATE TABLE dbo.irel_aree_localita (" + vbCrLf + _
				"	ral_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	ral_are_id int NOT NULL ," + vbCrLf + _
				"	ral_loc_id int NOT NULL " + vbCrLf + _
				" ); " + vbCrLf + _
		"CREATE TABLE dbo.irel_aree_comuni (" + vbCrLf + _
				"	rac_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rac_are_id int NOT NULL ," + vbCrLf + _
				"	rac_comune_codice nvarchar (6) NOT NULL " + vbCrLf + _
				" ); " + vbCrLf + _
		"ALTER TABLE dbo.irel_aree_comuni ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_aree_comuni_itb_aree" + vbCrLf + _
				"		FOREIGN KEY (rac_are_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_aree_comuni_tb_localita" + vbCrLf + _
				"		FOREIGN KEY (rac_comune_codice)" + vbCrLf + _
				"		REFERENCES dbo.tb_comuni ( Codice_ISTAT )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
		"ALTER TABLE dbo.irel_aree_localita ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_aree_localita_itb_aree" + vbCrLf + _
				"		FOREIGN KEY (ral_are_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_aree_localita_tb_localita" + vbCrLf + _
				"		FOREIGN KEY (ral_loc_id)" + vbCrLf + _
				"		REFERENCES dbo.tb_localita ( Loc_ID )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
	  " ALTER TABLE itb_eventi ADD"+ _
	  "		eve_bussola_it NTEXT NULL,"+ _
	  "		eve_bussola_en NTEXT NULL,"+ _
	  "		eve_bussola_fr NTEXT NULL,"+ _
	  "		eve_bussola_es NTEXT NULL,"+ _
	  "		eve_bussola_de NTEXT NULL; "
CALL DB.Execute(sql, 996)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 997
'...........................................................................................
sql = Aggiornamento__INFO__1(conn)
CALL DB.Execute(sql, 997)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 998
'...........................................................................................
sql = Aggiornamento__INFO__2(conn)
CALL DB.Execute(sql, 998)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 999
'...........................................................................................
sql = Aggiornamento__INFO__3(conn)
CALL DB.Execute(sql, 999)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1000
'...........................................................................................
sql = AggiornamentoSpeciale__INFO__4(DB, rs, 1000)
CALL DB.Execute(sql, 1000)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1000)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1001
'...........................................................................................
'aggiunge campo di visibilita' per gruppi di pubblicazione
'...........................................................................................
sql = " ALTER TABLE tb_pubblicazioni_APT ADD " + _
	  "		pub_gruppo_portale INT NULL  "
CALL DB.Execute(sql, 1001)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1002
'...........................................................................................
'aggiunge campi per gestione pubblicazione sul portale / APT
'su tabelle modelli e tipi strutture
'...........................................................................................
sql = "ALTER TABLE tb_modelli ADD " + _
	  "		mod_portale_pubblica BIT NULL " + _
	  " ; " + _
	  " ALTER TABLE tb_tipi_str ADD " + _
	  "		tip_portale_categoria INT NULL, " + _
	  "		temp_portale_nome_it nvarchar(250) NULL, " + _
	  "		temp_portale_nome_en nvarchar(250) NULL "
CALL DB.Execute(sql, 1002)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1003
'...........................................................................................
'crea relazioni non vincolanti tra gruppi di pubblicazione e gruppi descrittori
'e tra tipi di strutture e categorie di anagrafiche
'...........................................................................................
sql = " ALTER TABLE tb_tipi_str ADD CONSTRAINT FK_tb_tipi_str_itb_anagrafiche_tipi " + _
	  "		FOREIGN KEY ( tip_portale_categoria) " + _
	  "		REFERENCES itb_anagrafiche_tipi (ant_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_tipi_str NOCHECK CONSTRAINT FK_tb_tipi_str_itb_anagrafiche_tipi ; " + _
	  " ALTER TABLE tb_pubblicazioni_APT ADD CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag " + _
	  "		FOREIGN KEY ( pub_gruppo_portale ) " + _
	  "		REFERENCES itb_anagrafiche_descrRag (adr_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_pubblicazioni_APT NOCHECK CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag "
CALL DB.Execute(sql, 1003)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1004
'...........................................................................................
'imposta dati per sincronizzazione strutture
'...........................................................................................
sql = " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 18 " + vbCrLF + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Hotels' WHERE tip_mod_id=18 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 19 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Campeggi', temp_portale_nome_en='Campings' WHERE tip_mod_id=19 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 20 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Affittacamere', temp_portale_nome_en='Room rentals' WHERE tip_mod_id=20 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 21 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Appartamenti per vacanze', temp_portale_nome_en='Holiday apartments' WHERE tip_mod_id=21 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 22" + vbCrLf + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 23 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Guide turistiche' WHERE tip_mod_id=23 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 24 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Accompagnatori turistici' WHERE tip_mod_id=24 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 25 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Animatori turistici' WHERE tip_mod_id=25 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 26 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Guide naturalistico-ambientali' WHERE tip_mod_id=26 " + _
	  " ; " + _
 	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 27 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Residence' WHERE tip_mod_id=27 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 28 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Case per ferie', temp_portale_nome_en='Holiday houses' WHERE tip_id=32 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Ostelli', temp_portale_nome_en='Youth Hostels' WHERE tip_id=33 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Case religiose', temp_portale_nome_en='Religious guesthouses' WHERE tip_id=34 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Centri soggiorno-studi', temp_portale_nome_en='Schools with guestrooms' WHERE tip_id=35 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 29 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Bed & Breakfast' WHERE tip_mod_id=29 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 30" + vbCrLf + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 31 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Appartamenti per vacanze non classificati', temp_portale_nome_en='Ungraded holiday apartments' WHERE tip_mod_id=31 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 32 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Foresterie', temp_portale_nome_en='Guestrooms' WHERE tip_mod_id=32 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 33 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Country House' WHERE tip_mod_id=33 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 34 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Agenzie per affittanze', temp_portale_nome_en='Estate agencies for rental apartments' WHERE tip_mod_id=34 " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 35; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 36; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 37; " + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=1 WHERE mod_id = 38 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Agenzie di viaggio' WHERE tip_id=50 " + vbCrLf + _
	  " UPDATE tb_tipi_str SET temp_portale_nome_it='Associazioni ONLUS con attivita'' turistica' WHERE tip_id=51 " + vbCrLf + _
	  " ; " + _
	  " UPDATE tb_modelli SET mod_portale_pubblica=0 WHERE mod_id = 39; "
CALL DB.Execute(sql, 1004)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1005
'...........................................................................................
'rimuove relazioni perch&egrave; i dati vengono relazionati in modo non collegabile
'...........................................................................................
sql = " ALTER TABLE tb_tipi_str DROP CONSTRAINT FK_tb_tipi_str_itb_anagrafiche_tipi ; " + _
	  " ALTER TABLE tb_pubblicazioni_APT DROP CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag ; "
CALL DB.Execute(sql, 1005)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1006
'...........................................................................................
'corregge dati di sincronizzazione strutture
'rimuove relazioni perche' le unita' abitative verranno collegate come contatti interni
'e quindi non collegabili a rubriche
'...........................................................................................
sql = " UPDATE tb_indirizzario SET SyncroTable='VIEW_valid_strutture' WHERE SyncroTable LIKE 'VIEW_strutture' ; " + _
	  " UPDATE tb_rubriche SET SyncroTable='VIEW_valid_strutture' WHERE SyncroTable LIKE 'VIEW_strutture' ; " + _
	  " DELETE FROM rel_rub_ind WHERE id_indirizzo IN (SELECT IdElencoIndirizzi FROM tb_indirizzario WHERE IsNull(CntRel, 0)>0) ; " + _
	  " DELETE FROM tb_rubriche WHERE SyncroFilterTable LIKE 'tb_modelli' AND " + _
	  "		tb_rubriche.SyncroFilterKey IN (SELECT mod_id FROM tb_modelli WHERE mod_tipo_record='" & RECORD_TYPE_UA & "') "
CALL DB.Execute(sql, 1006)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1007
'...........................................................................................
'corregge dati pubblicazioni portali APT per dotazioni e servizi
'...........................................................................................
sql = " UPDATE tb_servizi SET serv_APT_nome_fra='', serv_APT_nome_spa='' ;" + _
	  " UPDATE tb_servizi SET serv_APT_nome_eng='', serv_APT_nome_fra='', serv_APT_nome_spa='', serv_APT_nome_ted='' " + _
	  " WHERE serv_ID IN (SELECT rel_grp_id_serv FROM rel_grp_serv INNER JOIN tb_pubblicazioni_apt " + _
	  					" ON rel_grp_serv.serv_APT_pubblicazione=tb_pubblicazioni_apt.pub_id " + _
						" WHERE pub_label_it LIKE '%guida%' AND pub_label_it LIKE '%disabili%') ; "	  
CALL DB.Execute(sql, 1007)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1008
'...........................................................................................
'aggiunge campo per relazione tra servizio e descrittore delle anagrafiche
'aggiunge relazioni tra tabelle del turismo e tabelle NEXT-INFO
'...........................................................................................
sql = " ALTER TABLE tb_servizi ADD " + _
	  "		serv_descrittore_portale INT NULL ; " + _
	  " ALTER TABLE tb_dotazioni ADD " + _
	  "		dotaz_descrittore_portale INT NULL ; " + _
	  " ALTER TABLE tb_servizi ADD CONSTRAINT FK_tb_servizi_itb_anagrafiche_descrittori " + _
	  "		FOREIGN KEY (serv_descrittore_portale) " + _
	  "		REFERENCES itb_anagrafiche_descrittori (and_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_servizi NOCHECK CONSTRAINT FK_tb_servizi_itb_anagrafiche_descrittori ; " + _
	  " ALTER TABLE tb_pubblicazioni_APT ADD CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag " + _
	  "		FOREIGN KEY ( pub_gruppo_portale ) " + _
	  "		REFERENCES itb_anagrafiche_descrRag (adr_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_pubblicazioni_APT NOCHECK CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag ; " + _
	  " ALTER TABLE tb_tipi_str ADD CONSTRAINT FK_tb_tipi_str_itb_anagrafiche_tipi " + _
	  "		FOREIGN KEY ( tip_portale_categoria) " + _
	  "		REFERENCES itb_anagrafiche_tipi (ant_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_tipi_str NOCHECK CONSTRAINT FK_tb_tipi_str_itb_anagrafiche_tipi ; " + _
	  " ALTER TABLE tb_dotazioni ADD CONSTRAINT FK_tb_dotazioni_itb_anagrafiche_descrittori " + _
	  "		FOREIGN KEY (dotaz_descrittore_portale) " + _
	  "		REFERENCES itb_anagrafiche_descrittori (and_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_dotazioni NOCHECK CONSTRAINT FK_tb_dotazioni_itb_anagrafiche_descrittori ; "	  
CALL DB.Execute(sql, 1008)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1009
'...........................................................................................
'sposta dati di pubblicazione servizi e dotazioni su tabella principale
'...........................................................................................
sql = " ALTER TABLE tb_servizi ADD " + _
	  "		serv_APT_pubblicazione INT NULL, " + _
	  "		serv_APT_ordine INT NULL " + _
	  " ; " + _
	  " UPDATE tb_servizi SET serv_APT_pubblicazione = (SELECT TOP 1 serv_APT_pubblicazione FROM rel_grp_serv WHERE rel_grp_id_serv=tb_servizi.serv_id), " + _
	  						" serv_APT_ordine = (SELECT TOP 1 serv_APT_ordine FROM rel_grp_serv WHERE rel_grp_id_serv=tb_servizi.serv_id ORDER BY serv_APT_ordine) " + _
	  " ; " + _
	  " ALTER TABLE tb_servizi ALTER COLUMN serv_APT_pubblicazione INT NOT NULL ; " + _
	  " ALTER TABLE tb_servizi ADD CONSTRAINT FK_tb_servizi__tb_pubblicazioni_APT " + _
	  "		FOREIGN KEY (serv_APT_pubblicazione) " + _
	  "		REFERENCES tb_pubblicazioni_APT(pub_id) " + _
	  "		ON DELETE NO ACTION " + _
	  "		ON UPDATE NO ACTION " + _
	  " ; " + _
	  " ALTER TABLE rel_grp_serv DROP CONSTRAINT FK_rel_grp_serv_tb_pubblicazioni_APT; " + _
	  " ALTER TABLE rel_Grp_serv DROP CONSTRAINT DF__rel_grp_s__serv___4F12BBB9 ; " + _
	  " ALTER TABLE rel_grp_serv DROP COLUMN serv_APT_pubblicazione ; " + _
	  " ALTER TABLE rel_grp_serv DROP COLUMN serv_APT_ordine ; " + _
	  " ALTER TABLE tb_dotazioni ADD " + _
	  "		dotaz_APT_pubblicazione INT NULL, " + _
	  "		dotaz_APT_ordine INT NULL " + _
	  " ; " + _
	  " UPDATE rel_grp_dotaz SET dotaz_APT_pubblicazione=9 WHERE rel_grp_id_dotaz=207; " + _
	  " UPDATE tb_dotazioni SET dotaz_APT_pubblicazione = (SELECT TOP 1 dotaz_APT_pubblicazione FROM rel_grp_dotaz WHERE rel_grp_id_dotaz = tb_dotazioni.dotaz_id), " + _
	                          " dotaz_APT_ordine = (SELECT TOP 1 dotaz_APT_ordine FROM rel_grp_dotaz WHERE rel_grp_id_dotaz = tb_dotazioni.dotaz_id ORDER BY dotaz_APT_ordine) " + _
	  " ; " + _
	  " ALTER TABLE tb_dotazioni ALTER COLUMN dotaz_APT_pubblicazione INT NOT NULL ; " + _
	  " ALTER TABLE tb_dotazioni ADD CONSTRAINT FK_tb_dotazioni__tb_pubblicazioni_APT " + _
	  "		FOREIGN KEY (dotaz_APT_pubblicazione ) " + _
	  "		REFERENCES tb_pubblicazioni_APT(pub_id) " + _
	  "		ON DELETE NO ACTION " + _
	  "		ON UPDATE NO ACTION " + _
	  "	; " + _
	  " ALTER TABLE rel_grp_dotaz DROP CONSTRAINT FK_rel_grp_dotaz_tb_pubblicazioni_APT ; " + _
	  " ALTER TABLE rel_Grp_dotaz DROP CONSTRAINT DF__rel_grp_d__dotaz__4D2A7347 ; " + _
	  " ALTER TABLE rel_grp_dotaz DROP COLUMN dotaz_APT_pubblicazione ; " + _
	  " ALTER TABLE rel_grp_dotaz DROP COLUMN dotaz_APT_ordine ; "
CALL DB.Execute(sql, 1009)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1010
'...........................................................................................
'rimuove pubblicazione della dotazione id: 328 ==> Fascia di eta' bambini
'...........................................................................................
sql = " UPDATE tb_dotazioni SET " + _
      " 	dotaz_APT_nome_ITA = '', " + _
	  " 	dotaz_APT_nome_ENG = '', " + _
	  " 	dotaz_APT_nome_FRA = '', " + _
	  " 	dotaz_APT_nome_TED = '', " + _
	  " 	dotaz_APT_nome_SPA = '', " + _
	  "		dotaz_APT_pubblicazione = 1, " + _
	  "		dotaz_APT_ordine = NULL " + _
	  " WHERE dotaz_id = 328 "
CALL DB.Execute(sql, 1010)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1011
'...........................................................................................
'rimozione tabelle, viste e stored per Vecchio WebService Hotel
'e vecchie stored procedure
'...........................................................................................
sql = DropObject(conn, "tb_wsusers", "TABLE") + _
	  DropObject(conn, "spWS_Hotel_Dotazioni", "PROCEDURE") + _
	  DropObject(conn, "spWS_Hotel_Header", "PROCEDURE") + _
	  DropObject(conn, "spWS_Hotel_List", "PROCEDURE") + _
	  DropObject(conn, "spWS_Hotel_Servizi", "PROCEDURE") + _
	  DropObject(conn, "spWS_Hotel_Updated", "PROCEDURE") + _
	  DropObject(conn, "spWS_Last_Update", "PROCEDURE") + _
	  DropObject(conn, "spWS_ListaComuni", "PROCEDURE") + _
	  DropObject(conn, "spWS_ListaLocalita", "PROCEDURE") + _
	  DropObject(conn, "spWS_ListaTipi", "PROCEDURE") + _
	  DropObject(conn, "Delete_Admin", "PROCEDURE") + _
	  DropObject(conn, "Delete_Sito", "PROCEDURE")
CALL DB.Execute(sql, 1011)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1012
'...........................................................................................
'modifica tipi associazioni di categoria per collegamento con categorie anagrafiche
'...........................................................................................
sql = " ALTER TABLE tb_tipiassoc ADD " + _
	  "		tipo_portale_categoria INT NULL ; " + _
	  " ALTER TABLE tb_tipiassoc ADD CONSTRAINT FK_tb_tipiassoc_itb_anagrafiche_tipi " + _
	  "		FOREIGN KEY (tipo_portale_categoria) " + _
	  "		REFERENCES itb_anagrafiche_tipi (ant_id) NOT FOR REPLICATION ; " + _
	  " ALTER TABLE tb_tipiassoc NOCHECK CONSTRAINT FK_tb_tipiassoc_itb_anagrafiche_tipi ; "
CALL DB.Execute(sql, 1012)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1013
'...........................................................................................
'aggiunge applicativo NEXT-WEB 5.0
'...........................................................................................
sql = " INSERT INTO tb_siti(id_sito, sito_nome, sito_amministrazione, sito_dir, sito_p1) " + _
      " VALUES (26, 'NEXT-web 5.0 [gestione grafica e contenuti accessibili]', 1, 'NextWeb5', 'WEB_ADMIN') ;"
CALL DB.Execute(sql, 1013)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1014
'...........................................................................................
'aggiunge colonne per collegamento tra categorie APT e nuove categorie APT
'...........................................................................................
sql = " ALTER TABLE itb_eventi_categorie ADD " + _
      "     evc_ListaCategorieApt_EXT ntext NULL; " + _
      " ALTER TABLE itb_anagrafiche_tipi ADD " + _
      "     ant_ListaCategorieApt_EXT ntext NULL; "
CALL DB.Execute(sql, 1014)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1015
'...........................................................................................
'importa dati delle categorie degli eventi e delle anagrafiche dal database di APPOGGIO
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1015)
if DB.last_update_executed then
    CALL Update_1015__import_CATEGORIE_NuovoPortale(iCatEventi, Application("DATA_import_categorie_ConnectionString"), 31)
    CALL Update_1015__import_CATEGORIE_NuovoPortale(iCatAnagrafiche, Application("DATA_import_categorie_ConnectionString"), 151)
    CALL Update_1015__import_CATEGORIE_NuovoPortale(iAree, Application("DATA_import_categorie_ConnectionString"), 0)
end if

'funzione di import delle categorie di default dalla connessione indicata (DATABASE DI APPOGGIO)
sub Update_1015__import_CATEGORIE_NuovoPortale(oCategorie, ReadConnectionString, MaxId)
	dim connSource
    Set connSource = Server.CreateObject("ADODB.connection")
	connSource.open ReadConnectionString
    
    CALL CopyTableData(connSource, conn, _
                       oCategorie.tabella & IIF(MaxId>0, " WHERE " & oCategorie.prefisso & "_id<=" & MaxId, "") & " ORDER BY " & oCategorie.prefisso & "_livello, " & oCategorie.prefisso & "_id", _
                       oCategorie.tabella, oCategorie.prefisso & "_id", true)

    connSource.close
    set connSource = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1016
'...........................................................................................
'Genera le categorie corrispondenti ai tipi di strutture
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1016)
if DB.last_update_executed then
	CALL Aggiornamento_1016_GenerazioneCategorieModelli(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1016_GenerazioneCategorieModelli(conn, rs)
	dim sql, rsT, rsC, CategoriaPadre, DesList
    set rsT = Server.CreateObject("ADODB.Recordset")
    set rsC = Server.CreateObject("ADODB.Recordset")
    
    'importa descrittori comuni
    
    DesList = import_DescrittoreAnagrafiche(conn, rs, "VIEW_valid_strutture", "open_1", _
						          	        "Periodo di apertura - 1", "", "", "", "", _
									        "", adVarChar, "", NULL)
	DesList = DesList & ", " & import_DescrittoreAnagrafiche(conn, rs, "VIEW_valid_strutture", "open_2", _
						          	                         "Periodo di apertura - 2", "", "", "", "", _
                                                             "", adVarChar, "", NULL)
	DesList = DesList & ", " & import_DescrittoreAnagrafiche(conn, rs, "VIEW_valid_strutture", "open_3", _
						          	                         "Periodo di apertura - 3", "", "", "", "", _
                                                             "", adVarChar, "", NULL)
	DesList = DesList & ", " & import_DescrittoreAnagrafiche(conn, rs, "VIEW_valid_strutture", "open_4", _
                                                             "Periodo di apertura - 4", "", "", "", "", _
                                                             "", adVarChar, "", NULL)
    
    'lista tipologie da importare.
	sql = " SELECT * FROM tb_tipi_str INNER JOIN tb_modelli ON tb_tipi_str.tip_mod_id = tb_modelli.mod_id " + _
	      " WHERE mod_portale_pubblica=1 " + _
		  " ORDER BY mod_id, temp_portale_nome_it "
	rsT.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	while not rsT.eof
        
        'recupera categorie padre
	    CategoriaPadre = import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, GetCodice(CODE_ASSESSORATO, rst("mod_tipo_record"), ""))
        sql = "SELECT * FROM itb_anagrafiche_tipi WHERE ant_id=" & CategoriaPadre
        CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rsc, sql, DesList, false)
        
        'recupera categoria per la tipologia in corso
        sql = "SELECT * FROM itb_anagrafiche_tipi WHERE ant_padre_id=" & CategoriaPadre & " AND " & _
              " ant_nome_it LIKE '" & ParseSql(rsT("temp_portale_nome_it"), adChar) & "' "
        rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        
        rsT("tip_portale_categoria") = import_syncro_CATEGORIA(conn, rs, iCatAnagrafiche, CategoriaPadre, GetCodice(CODE_ASSESSORATO, rst("mod_tipo_record"), rsT("tip_id")), "tb_tipi_str", rsT("tip_id"), rsT("temp_portale_nome_it"), rsT("temp_portale_nome_en"), "", "", "", rsT("tip_id"))
        
        sql = "SELECT * FROM itb_anagrafiche_tipi WHERE ant_id=" & rsT("tip_portale_categoria")
        CALL import_Associa_DescrittoriCategorieAnagrafiche(conn, rsc, sql, DesList, false)
        
        rsT.update
		rsT.movenext
	wend

	rsT.close
    set rsT = nothing
    set rsC = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1017
'...........................................................................................
'Genera i gruppi di descrittori corrispondenti ai gruppi di pubblicazioni APT
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1017)
if DB.last_update_executed then
	CALL Aggiornamento_1017_GenerazioneGruppiPubblicazione(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1017_GenerazioneGruppiPubblicazione(conn, rs, rst)
	dim sql
	sql = "SELECT * FROM tb_pubblicazioni_APT WHERE pub_order>0"
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		sql = " UPDATE tb_pubblicazioni_APT SET pub_gruppo_portale= " & _
		      import_GruppoDescrittoreAnagrafiche(conn, rst, "tb_pubblicazioni_APT", rs("pub_id"), rs("pub_label_it"), rs("pub_label_en"), "", "", "", rs("pub_order")) & _
			  " WHERE pub_id=" & rs("pub_id")
		CALL conn.execute(sql, ,adExecuteNoRecords)
		rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1018
'...........................................................................................
'modifica dei dati di sincronizzazione delle strutture ricettive nel next-Com
'aggiorna unit&agrave; abitative portandole a contatto interno del proprietario
'...........................................................................................
sql = " UPDATE tb_indirizzario " + _
      " SET tb_indirizzario.CntRel = (SELECT IdElencoIndirizzi FROM tb_indirizzario WHERE SyncroKey LIKE RTRIM(IsNull(Cod_proprietario, '')) ) " + _
      " FROM tb_indirizzario INNER JOIN VIEW_strutture ON tb_indirizzario.SyncroKey LIKE RTRIM(VIEW_strutture.RegCode); " + _
      " DELETE FROM rel_rub_ind WHERE id_indirizzo IN (SELECT IdElencoIndirizzi FROM tb_indirizzario WHERE IsNull(cntRel, 0)>0) "
CALL DB.Execute(sql, 1018)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1019
'...........................................................................................
'Import descrittori da servizi pubblicati nelle APT
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1019)
if DB.last_update_executed then
	CALL Aggiornamento_1019_ImportServizi_Descrittori(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1019_ImportServizi_Descrittori(conn, rs, rst)
	dim sql, DesTipo
	sql = "SELECT * FROM tb_servizi INNER JOIN tb_pubblicazioni_APT " + _
		  " ON tb_servizi.serv_APT_pubblicazione=tb_pubblicazioni_apt.pub_id " + _
		  " WHERE IsNull(pub_gruppo_portale,0) > 0 " 
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		if not rs("serv_val") then
			DesTipo = adBoolean
		else
			Select case rs("serv_val_typ")
				case TYPE_TESTO
					DesTipo = adVarChar
				case TYPE_NUMERICO
					DesTipo = adNumeric
			end select
		end if

		'inserisce e sincronizza descrittore
		rs("serv_descrittore_portale") = import_DescrittoreAnagrafiche(conn, rst, "tb_servizi", rs("serv_id"), _
										  							   rs("serv_APT_nome_ITA"), _
																	   IIF(rs("serv_APT_nome_ENG")<>rs("serv_APT_nome_ITA"),rs("serv_APT_nome_ENG"), "") , _
																	   IIF(rs("serv_APT_nome_FRA")<>rs("serv_APT_nome_ITA"),rs("serv_APT_nome_FRA"), "") , _
																	   IIF(rs("serv_APT_nome_TED")<>rs("serv_APT_nome_ITA"),rs("serv_APT_nome_TED"), "") , _
																	   IIF(rs("serv_APT_nome_SPA")<>rs("serv_APT_nome_ITA"),rs("serv_APT_nome_SPA"), "") , _
											  						   "", DesTipo, rs("serv_symb"), rs("pub_gruppo_portale"))
		rs.update
		rs.movenext
	wend

	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1020
'...........................................................................................
'Import descrittori da dotazioni pubblicati nelle APT
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1020)
if DB.last_update_executed then
	CALL Aggiornamento_1020_ImportDotazioni_Descrittori(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1020_ImportDotazioni_Descrittori(conn, rs, rst)
	dim sql, DesTipo
	sql = " SELECT * FROM tb_dotazioni INNER JOIN tb_pubblicazioni_APT " + _
		  " ON tb_dotazioni.dotaz_APT_pubblicazione = tb_pubblicazioni_APT.pub_id " + _
		  " WHERE IsNull(pub_gruppo_portale,0)>0"
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		DesTipo = NULL
		Select Case rs("dotaz_num_val")
			case 1
				Select Case rs("dotaz_typ")
					case TYPE_NUMERICO
						DesTipo = adNumeric
					case TYPE_TESTO
						DesTipo = adVarChar
				end select
			case 2
				Select Case rs("dotaz_typ")
					case TYPE_NUMERICO, TYPE_NUMERICO_LINEA
						DesTipo = adNumeric
					case TYPE_TESTO, TYPE_TESTO_LINEA, TYPE_SCELTA
						DesTipo = adVarChar
				end select
			case 3
				Select Case rs("dotaz_typ")
					case TYPE_NUMERICO, TYPE_NUMERICO_LINEA
						DesTipo = adNumeric
				end select
			case 4
				Select Case rs("dotaz_typ")
					case TYPE_NUMERICO
						DesTipo = adNumeric
					case TYPE_PREZZO
						DesTipo = adDouble
				end select
			case 5
				Select Case rs("dotaz_typ")
					case TYPE_NUMERICO
						DesTipo = adNumeric
					case TYPE_FLAG
						DesTipo = adDouble
				end select
			case 6
				Select Case rs("dotaz_typ")
					case TYPE_FLAG
						DesTipo = adDouble
				end select
		end select
		
		'inserisce e sincronizza descrittore
		rs("dotaz_descrittore_portale") = import_DescrittoreAnagrafiche(conn, rst, "tb_dotazioni", rs("dotaz_id"), _
										  							    rs("dotaz_APT_nome_ITA"), _
																		IIF(rs("dotaz_APT_nome_ENG")<>rs("dotaz_APT_nome_ITA"),rs("dotaz_APT_nome_ENG"), "") , _
																		IIF(rs("dotaz_APT_nome_FRA")<>rs("dotaz_APT_nome_ITA"),rs("dotaz_APT_nome_FRA"), "") , _
																		IIF(rs("dotaz_APT_nome_TED")<>rs("dotaz_APT_nome_ITA"),rs("dotaz_APT_nome_TED"), "") , _
																		IIF(rs("dotaz_APT_nome_SPA")<>rs("dotaz_APT_nome_ITA"),rs("dotaz_APT_nome_SPA"), "") , _
											  						    "", DesTipo, rs("dotaz_symb"), rs("pub_gruppo_portale"))
		rs.update
		rs.movenext
	wend

	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1021
'...........................................................................................
'Associa descrittori dei servizi alle categorie di anagrafiche
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1021)
if DB.last_update_executed then
	CALL Aggiornamento_1021_Associazione__Servizi_Descrittori__Categorie_Anagrafiche(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1021_Associazione__Servizi_Descrittori__Categorie_Anagrafiche(conn, rs, rst)
	dim sql
    
	sql = " SELECT tip_portale_categoria, serv_descrittore_portale, rel_grp_serv_ord FROM " + _
		  " tb_modelli INNER JOIN tb_tipi_str ON tb_modelli.mod_id = tb_tipi_str.tip_mod_id " + _
		  " INNER JOIN tb_grp_vis ON tb_modelli.mod_id = tb_grp_vis.grp_mod_id " + _
		  " INNER JOIN rel_grp_serv ON tb_grp_vis.grp_id = rel_grp_serv.rel_id_grp_serv " + _
		  " INNER JOIN tb_servizi ON rel_grp_serv.rel_grp_id_serv = tb_servizi.serv_id " + _
		  " WHERE IsNull(tip_portale_categoria, 0)>0 AND IsNull(serv_descrittore_portale, 0)>0 "
	rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	while not rs.eof
        CALL Syncro_AssociazioneDescrittoriCategorieAnagrafiche(conn, rst, rs("tip_portale_categoria"), rs("serv_descrittore_portale"), cInteger(rs("rel_grp_serv_ord")), true)
        rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1022
'...........................................................................................
'Associa descrittori delle dotazioni alle categorie di anagrafiche
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1022)
if DB.last_update_executed then
	CALL Aggiornamento_1022_Associazione__dotazioni_Descrittori__Categorie_Anagrafiche(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1022_Associazione__dotazioni_Descrittori__Categorie_Anagrafiche(conn, rs, rst)
	dim sql   
	sql = " SELECT tip_portale_categoria, dotaz_descrittore_portale, rel_grp_dotaz_ordine FROM " + _
		  " tb_modelli INNER JOIN tb_tipi_str ON tb_modelli.mod_id = tb_tipi_str.tip_mod_id " + _
		  " INNER JOIN tb_grp_vis ON tb_modelli.mod_id = tb_grp_vis.grp_mod_id " + _
		  " INNER JOIN rel_grp_dotaz ON tb_grp_vis.grp_id = rel_grp_dotaz.rel_id_grp_dotaz " + _
		  " INNER JOIN tb_dotazioni ON rel_grp_dotaz.rel_grp_id_dotaz = tb_dotazioni.dotaz_id " + _
		  " WHERE IsNull(tip_portale_categoria, 0)>0 AND IsNull(dotaz_descrittore_portale, 0)>0 "
	rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	while not rs.eof
        CALL Syncro_AssociazioneDescrittoriCategorieAnagrafiche(conn, rst, rs("tip_portale_categoria"), rs("dotaz_descrittore_portale"), cInteger(rs("rel_grp_dotaz_ordine")), true)
        rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1023
'...........................................................................................
'rimozione colonne di ordinamento delle dotazioni e servizi in pubblicazione APT
'non piu' necessari perche' ora gestiti dai descrittori
'...........................................................................................
sql = " ALTER TABLE tb_servizi DROP COLUMN serv_APT_ordine; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_nome_ITA; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_nome_ENG; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_nome_FRA; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_nome_TED; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_nome_SPA; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_ordine; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_nome_ITA; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_nome_ENG; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_nome_FRA; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_nome_TED; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_nome_SPA; "
CALL DB.Execute(sql, 1023)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1024
'...........................................................................................
'rimozione gruppi di pubblicazione nei portali della APT, relazioni e chiavi esterne di
'dotazioni e servizi; sblocca anche raggruppamenti di descrittori collegati alle APT
'...........................................................................................
sql = " ALTER TABLE dbo.tb_pubblicazioni_APT DROP CONSTRAINT FK_tb_pubblicazioni_APT_itb_anagrafiche_descrRag ; " + _
	  " ALTER TABLE dbo.tb_servizi DROP CONSTRAINT FK_tb_servizi__tb_pubblicazioni_APT ; " + _
	  " ALTER TABLE dbo.tb_dotazioni DROP CONSTRAINT FK_tb_dotazioni__tb_pubblicazioni_APT ; " + _
	  " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_APT_pubblicazione; " + _
	  " ALTER TABLE tb_servizi DROP COLUMN serv_APT_pubblicazione ; " + _
	  DropObject(conn, "tb_pubblicazioni_APT", "TABLE") + _
	  " UPDATE itb_anagrafiche_DescrRag SET adr_external_id=NULL,  adr_external_source=NULL WHERE adr_external_source LIKE 'tb_pubblicazioni_apt' "
CALL DB.Execute(sql, 1024)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1024)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'INIZIO DELL'IMPORT DELLE CATEGORIE DEI DATI PROVENIENTI DALLE APT
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1025
'...........................................................................................
'Import delle categorie di Eventi
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1025)
if DB.last_update_executed then
    CALL import_CATEGORIE_APT(iCatEventi, CODE_APTBC, CODE_EVENTI, CODE_EVENTI, "CategorieEventi", "IDC", "Descrizione", "Desc_eng", "Desc_fra", "Desc_ted", "Desc_spa")
    CALL import_CATEGORIE_APT(iCatEventi, CODE_APTJE, CODE_EVENTI, CODE_EVENTI, "CategorieEventi", "IDC", "Descrizione", "Desc_eng", "Desc_fra", "Desc_ted", "Desc_spa")
    CALL import_CATEGORIE_APT(iCatEventi, CODE_APTVE, CODE_EVENTI, CODE_EVENTI, "CategorieEventi", "IDC", "Descrizione", "Desc_eng", "Desc_fra", "Desc_ted", "Desc_spa")
    CALL import_CATEGORIE_APT(iCatEventi, CODE_APTCH, CODE_EVENTI, CODE_EVENTI, "CategorieEventi", "IDC", "Descrizione", "Desc_eng", "Desc_fra", "Desc_ted", "Desc_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1026
'...........................................................................................
'Import delle aree
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1026)
if DB.last_update_executed then
    CALL import_AREE(iAree, CODE_APTBC)
    CALL import_AREE(iAree, CODE_APTJE)
    CALL import_AREE(iAree, CODE_APTVE)
    CALL import_AREE(iAree, CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1027
'...........................................................................................
'Import delle categorie di Luoghi
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1027)
if DB.last_update_executed then
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTBC, CODE_LUOGHI, CODE_LUOGHI, "TipoLuoghi", "IDL", "Tipo_luogo", "Tipo_luogo_eng", "Tipo_luogo_fra", "Tipo_luogo_ted", "Tipo_luogo_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTJE, CODE_LUOGHI, CODE_LUOGHI, "TipoLuoghi", "IDL", "Tipo_luogo", "Tipo_luogo_eng", "Tipo_luogo_fra", "Tipo_luogo_ted", "Tipo_luogo_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTVE, CODE_LUOGHI, CODE_LUOGHI, "TipoLuoghi", "IDL", "Tipo_luogo", "Tipo_luogo_eng", "Tipo_luogo_fra", "Tipo_luogo_ted", "Tipo_luogo_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTCH, CODE_LUOGHI, CODE_LUOGHI, "TipoLuoghi", "IDL", "Tipo_luogo", "Tipo_luogo_eng", "Tipo_luogo_fra", "Tipo_luogo_ted", "Tipo_luogo_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1028
'...........................................................................................
'Import dei tipi di notizie utili
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1028)
if DB.last_update_executed then
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTBC, CODE_NOTIZIE_T, CODE_NOTIZIE, "Tipi_notutil", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTJE, CODE_NOTIZIE_T, CODE_NOTIZIE, "Tipi_notutil", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTVE, CODE_NOTIZIE_T, CODE_NOTIZIE, "Tipi_notutil", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTCH, CODE_NOTIZIE_T, CODE_NOTIZIE, "Tipi_notutil", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1029
'...........................................................................................
'Import dei sottotipi di notizie utili
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1029)
if DB.last_update_executed then
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTBC, CODE_NOTIZIE, CODE_NOTIZIE_T, "SottoTipi_Notutil", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTJE, CODE_NOTIZIE, CODE_NOTIZIE_T, "SottoTipi_Notutil", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTVE, CODE_NOTIZIE, CODE_NOTIZIE_T, "SottoTipi_Notutil", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTCH, CODE_NOTIZIE, CODE_NOTIZIE_T, "SottoTipi_Notutil", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1030
'...........................................................................................
'Import dei tipi di locali e servizi
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1030)
if DB.last_update_executed then
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTBC, CODE_LOCALI_T, CODE_LOCALI, "Tipi_LS", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTJE, CODE_LOCALI_T, CODE_LOCALI, "Tipi_LS", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTVE, CODE_LOCALI_T, CODE_LOCALI, "Tipi_LS", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
    CALL import_CATEGORIE_APT(iCatAnagrafiche, CODE_APTCH, CODE_LOCALI_T, CODE_LOCALI, "Tipi_LS", "id_tipoutil", "tipo_nome_it", "tipo_nome_eng", "tipo_nome_fra", "tipo_nome_ted", "tipo_nome_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1031
'...........................................................................................
'Import dei sottotipi di locali e servizi
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1031)
if DB.last_update_executed then
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTBC, CODE_LOCALI, CODE_LOCALI_T, "SottoTipi_LocServ", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTJE, CODE_LOCALI, CODE_LOCALI_T, "SottoTipi_LocServ", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTVE, CODE_LOCALI, CODE_LOCALI_T, "SottoTipi_LocServ", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
    CALL import_SOTTO_CATEGORIE_APT(iCatAnagrafiche, CODE_APTCH, CODE_LOCALI, CODE_LOCALI_T, "SottoTipi_LocServ", "id_sottipo", "ref_tipo", "sottip_nome_it", "sottip_nome_eng", "sottip_nome_fra", "sottip_nome_ted", "sottip_nome_spa")
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'FINE IMPORT DATI DELLE CATEGORIE
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************


'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'Aggiorna tutte le anagrafiche del NEXT-INFO con sincronizzazione definitiva
'l'aggiornamento e' stato spezzato in esecuzioni multiple per problemi di esecuzione
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1032
'IMPORT ALBERGHI di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1032)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 18, "04")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1033
'IMPORT ALBERGHI di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1033)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 18, "05")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1034
'IMPORT ALBERGHI di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1034)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 18, "06")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1035
'IMPORT ALBERGHI di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1035)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 18, "07")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1036
'IMPORT Campeggi
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1036)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 19, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1037
'IMPORT Affittacamere
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1037)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 20, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1038
'IMPORT u.a. classificate
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1038)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 21, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1039
'IMPORT proprietari u.a. classificate
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1039)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 22, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1040
'IMPORT guide turistiche
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1040)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 23, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1041
'IMPORT accompagnatori turistici
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1041)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 24, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1042
'IMPORT animatori turistici
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1042)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 25, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1043
'IMPORT guide naturalistiche
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1043)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 26, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1044
'IMPORT residence
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1044)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 27, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1045
'IMPORT ricettivita' sociali
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1045)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 28, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1046
'IMPORT bed & breakfast
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1046)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 29, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1047
'IMPORT proprietari u.a. non classificate di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1047)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 30, "04")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1048
'IMPORT proprietari u.a. non classificate di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1048)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 30, "05")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1049
'IMPORT proprietari u.a. non classificate di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1049)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 30, "06")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1050
'IMPORT proprietari u.a. non classificate di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1050)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 30, "07")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1051
'IMPORT u.a. non classificate di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1051)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 31, "04")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1052
'IMPORT u.a. non classificate di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1052)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 31, "05")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1053
'IMPORT u.a. non classificate di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1053)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 31, "06")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1054
'IMPORT u.a. non classificate di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1054)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 31, "07")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1055
'IMPORT foresterie
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1055)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 32, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1056
'IMPORT country house
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1056)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 33, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1057
'IMPORT Agenzie immobiliari u.a.
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1057)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 34, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1058
'IMPORT u.a. gestite da agenzie immobiliari di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1058)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 36, "04")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1059
'IMPORT u.a. gestite da agenzie immobiliari di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1059)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 36, "05")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1060
'IMPORT u.a. gestite da agenzie immobiliari di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1060)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 36, "06")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1061
'IMPORT u.a. gestite da agenzie immobiliari di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1061)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 36, "07")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1062
'IMPORT direttori tecnici
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1062)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 37, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1063
'IMPORT agenzie di viaggio e turismo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1063)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 38, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1064
'IMPORT accompagnatori agenzie di viaggio
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1064)
if DB.last_update_executed then
	CALL Aggiornamento_multiplo_Dati_NEXTINFO(DB.objConn, rs, rst, 39, "")
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_multiplo_Dati_NEXTINFO(conn, rs, rst, MODELLO, AptCode)
	dim sql, readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.open conn.ConnectionString, "", ""
	readconn.CommandTimeout = 180
	
	sql = " SELECT modello, RegCode, mod_tipo_record FROM VIEW_Strutture " + _
		  " WHERE modello = " & modello & _
		  IIF(AptCode <> "", " AND AptCode LIKE '%" & AptCode & "%' ", "")
	readRs.open sql, readConn, adOpenStatic, adLockReadOnly, adCmdText
	while not readRs.eof %>
		<!-- <%= readRs("RegCode") %> - <%= readRs.absoluteposition %> su <%= readRs.recordcount %>-->
		<%CALL SincronizzaStruttura_NextCom_NextInfo(readConn, Conn, rs, readRs("modello"), readRs("RegCode"), readRs("mod_tipo_record"), true)
		readRs.movenext
	wend
	readRs.close
	
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1065
'...........................................................................................
'genera le categorie corrispondenti alle categorie di associazioni di categoria
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1065)
if DB.last_update_executed then
	CALL Aggiornamento_1065_GenerazioneCategorieAssociazioni(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1065_GenerazioneCategorieAssociazioni(conn, rs, rst)
	dim sql, CategoriaPadre
	
	'recupera dati categoria di base per le associazioni
	CategoriaPadre = import_GetCategoriaPrincipale(iCatAnagrafiche, conn, rs, GetCodice(CODE_ASSESSORATO, CODE_ASSOCIAZIONI, ""))
    
    sql = "SELECT * FROM tb_tipiassoc "
	set rst = server.createObject("ADODB.recordset")
	rsT.open sql, conn, adOpenStatic, adLockOptimistic
	while not rsT.eof
        
        'recupera categoria per la tipologia in corso
        sql = "SELECT * FROM itb_anagrafiche_tipi WHERE ant_padre_id=" & CategoriaPadre & " AND " & _
              " ant_nome_it LIKE '" & ParseSql(rsT("nome_tipo"), adChar) & "' "
        rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
        
        rsT("tipo_portale_categoria") = import_syncro_CATEGORIA(conn, rs, iCatAnagrafiche, CategoriaPadre, GetCodice(CODE_ASSESSORATO, CODE_ASSOCIAZIONI, rsT("id_tipo")), "tb_tipiassoc", rsT("id_tipo"), rsT("nome_tipo"), "", "", "", "", rsT("id_tipo"))
        rsT.update
        
		rsT.movenext
	wend
	rsT.close
	
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1066
'...........................................................................................
'Importa le associazioni di categoria nelle anagrafiche
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1066)
if DB.last_update_executed then
	CALL Aggiornamento_1066_import_associazioni_categoria(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1066_import_associazioni_categoria(conn, rs, rst)
	dim sql, readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.Open Application(request("ConnString")), "", ""
	readconn.CommandTimeout = 180
	
	sql = " SELECT * FROM tb_assoc"
	readRs.open sql, readConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	while not readRs.eof %>
		<!-- <%= readRs("asc_id") %> -->
		<%CALL SincronizzaAssociazione_NextCom_NextInfo(Conn, rs, readRs("asc_id"))
		readRs.movenext
	wend
	readRs.close
	
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1067
'...........................................................................................
'rimuove campi temporanei su tipi strutture
'...........................................................................................
sql = " ALTER TABLE tb_tipi_str DROP COLUMN temp_portale_nome_it, temp_portale_nome_en "
CALL DB.Execute(sql, 1067)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'INIZIO DELL'IMPORT DEI DATI PROVENIENTI DALLE APT
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1068
'...........................................................................................
'Impostazione dei descrittori delle anagrafiche e relative associazioni con le categorie interessate
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1068)
if DB.last_update_executed then
    CALL import_DESCRITTORI_ANAGRAFICHE__MAPPE()
    CALL import_DESCRITTORI_ANAGRAFICHE__SPIAGGE()
    CALL import_DESCRITTORI_ANAGRAFICHE__LUOGHI()
    CALL import_DESCRITTORI_ANAGRAFICHE__LOCALI_E_SERVIZI()
    CALL import_DESCRITTORI_ANAGRAFICHE__STRUTTURE_RICETTIVE_NON_SYNCRO()
end if
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati delle spiagge
'...........................................................................................
'AGGIORNAMENTO 1069
'import spiagge apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1069)
if DB.last_update_executed then
    CALL import_SPIAGGE(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1070
'import spiagge apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1070)
if DB.last_update_executed then
    CALL import_SPIAGGE(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1071
'import spiagge apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1071)
if DB.last_update_executed then
    CALL import_SPIAGGE(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1072
'import spiagge apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1072)
if DB.last_update_executed then
    CALL import_SPIAGGE(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati dei luoghi
'...........................................................................................
'AGGIORNAMENTO 1073
'import luoghi apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1073)
if DB.last_update_executed then
    CALL import_LUOGHI(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1074
'import luoghi apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1074)
if DB.last_update_executed then
    CALL import_LUOGHI(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1075
'import luoghi apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1075)
if DB.last_update_executed then
    CALL import_LUOGHI(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1076
'import luoghi apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1076)
if DB.last_update_executed then
    CALL import_LUOGHI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1077
'...........................................................................................
'Import degli eventi speciali dai portali delle APT
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1077)
if DB.last_update_executed then
    CALL import_EVENTI_SPECIALI(CODE_APTBC)
    CALL import_EVENTI_SPECIALI(CODE_APTJE)
    CALL import_EVENTI_SPECIALI(CODE_APTVE)
    CALL import_EVENTI_SPECIALI(CODE_APTCH)
end if
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati degli eventi
'...........................................................................................
'AGGIORNAMENTO 1078
'import eventi apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1078)
if DB.last_update_executed then
    CALL import_EVENTI(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1079
'import eventi apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1079)
if DB.last_update_executed then
    CALL import_EVENTI(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1080
'import eventi apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1080)
if DB.last_update_executed then
    CALL import_EVENTI(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1081
'import eventi apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1081)
if DB.last_update_executed then
    CALL import_EVENTI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati delle notizie utili
'...........................................................................................
'AGGIORNAMENTO 1082
'import notizie utili apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1082)
if DB.last_update_executed then
    CALL import_NOTIZIE_UTILI(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1083
'import notizie utili apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1083)
if DB.last_update_executed then
    CALL import_NOTIZIE_UTILI(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1084
'import notizie utili apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1084)
if DB.last_update_executed then
    CALL import_NOTIZIE_UTILI(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1085
'import notizie utili apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1085)
if DB.last_update_executed then
    CALL import_NOTIZIE_UTILI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati dei locali e servizi
'...........................................................................................
'AGGIORNAMENTO 1086
'import locali e servizi apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1086)
if DB.last_update_executed then
    CALL import_LOCALI_E_SERVIZI(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1087
'import locali e servizi apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1087)
if DB.last_update_executed then
    CALL import_LOCALI_E_SERVIZI(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1088
'import locali e servizi apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1088)
if DB.last_update_executed then
    CALL import_LOCALI_E_SERVIZI(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1089
'import locali e servizi apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1089)
if DB.last_update_executed then
    CALL import_LOCALI_E_SERVIZI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati delle strutture ricettive non sincronizzate
'...........................................................................................
'AGGIORNAMENTO 1090
'import strutture ricettive non sincronizzate apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1090)
if DB.last_update_executed then
    CALL import_STRUTTURE_Non_Sincronizzate(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1091
'import strutture ricettive non sincronizzate apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1091)
if DB.last_update_executed then
    CALL import_STRUTTURE_Non_Sincronizzate(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1092
'import strutture ricettive non sincronizzate apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1092)
if DB.last_update_executed then
    CALL import_STRUTTURE_Non_Sincronizzate(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1093
'import strutture ricettive non sincronizzate apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1093)
if DB.last_update_executed then
    CALL import_STRUTTURE_Non_Sincronizzate(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'IMPORT dati delle strutture ricettive sincronizzate
'...........................................................................................
'AGGIORNAMENTO 1094
'import strutture ricettive sincronizzate apt di Bibione e Caorle
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1094)
if DB.last_update_executed then
    CALL import_STRUTTURE_SINCRONIZZATE(CODE_APTBC)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1095
'import strutture ricettive sincronizzate apt di Jesolo
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1095)
if DB.last_update_executed then
    CALL import_STRUTTURE_SINCRONIZZATE(CODE_APTJE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1096
'import strutture ricettive sincronizzate apt di Venezia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1096)
if DB.last_update_executed then
    CALL import_STRUTTURE_SINCRONIZZATE(CODE_APTVE)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'AGGIORNAMENTO 1097
'import strutture ricettive sincronizzate apt di Chioggia
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1097)
if DB.last_update_executed then
    CALL import_STRUTTURE_SINCRONIZZATE(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'FINE IMPORT DATI PROVENIENTI DALLE APT
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1098
'...........................................................................................
sql = Aggiornamento__INFO__5(conn)
CALL DB.Execute(sql, 1098)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1099
'...........................................................................................
'Importa le tabelle del sito dal database di import
'...........................................................................................
sql = " DELETE FROM ptb_categorieGallery; "
CALL DB.Execute(sql, 1099)
if DB.last_update_executed then
	CALL Aggiornamento_1099_import_sito(DB.objConn)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1099_import_sito(Conn)
	dim SourceConn
	set SourceConn = Server.CreateObject("ADODB.Connection")
	SourceConn.Open Application("DATA_import_sito_ConnectionString"), "", ""
	
    'importa dati del NEXT-PASSPORT
    CALL CopyTableData(SourceConn, Conn, "tb_siti", "", "id_sito", false)
    CALL CopyTableData(SourceConn, Conn, "tb_siti_parametri WHERE NOT (par_key LIKE 'RUBRICA_ANAGRAFICHE')", "tb_siti_parametri", "par_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_siti_tabelle", "", "tab_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_siti_tabelle_pubblicazioni", "", "pub_id", true)
    
    'importa dati NEXT-WEB5
    CALL CopyTableData(SourceConn, Conn, "tb_webs", "", "id_webs", false)
    CALL CopyTableData(SourceConn, Conn, "tb_menu", "", "m_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_menuItem", "", "mi_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_css_groups", "", "grp_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_css_styles", "", "style_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_objects", "", "id_objects", true)
    CALL CopyTableData(SourceConn, Conn, "tb_tipo", "", "id_tip", false)
    CALL CopyTableData(SourceConn, Conn, "tb_paginesito", "", "id_pagineSito", true)
    CALL CopyTableData(SourceConn, Conn, "tb_pages", "", "id_page", true)
    CALL CopyTableData(SourceConn, Conn, "tb_layers", "", "id_lay", true)
    
    'importa dati index-content
    CALL CopyTableData(SourceConn, Conn, "tb_contents WHERE co_id IN (SELECT co_id FROM v_indice WHERE tab_sito_id <> 29)", "tb_contents", "co_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_contents_index WHERE idx_content_id IN (SELECT co_id FROM v_indice WHERE tab_sito_id <> 29) ORDER BY idx_livello, idx_id", "tb_contents_index", "idx_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_siti_tabelle_pubblicazioni", "", "pub_id", true)
    CALL CopyTableData(SourceConn, Conn, "rel_index_pubblicazioni WHERE rip_idx_id IN (SELECT idx_id FROM v_indice WHERE tab_sito_id <> 29)", "rel_index_pubblicazioni", "rip_id", true)
    CALL CopyTableData(SourceConn, Conn, "rel_index_admin", "", "ria_id", true)
    
    'importa dati next-link
    CALL CopyTableData(SourceConn, Conn, "tb_links", "", "link_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_links_categorie", "", "cat_id", true)
    
    'importa dati next-faq
    CALL CopyTableData(SourceConn, Conn, "tb_FAQ", "", "faq_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_FAQ_categorie", "", "cat_id", true)
    
    'importa dati next-news
    CALL CopyTableData(SourceConn, Conn, "tb_news", "", "news_id", true)
    CALL CopyTableData(SourceConn, Conn, "tb_news_categorie", "", "cat_id", true)
    
    'importa dati next-gallery
    CALL CopyTableData(SourceConn, Conn, "ptb_categorieGallery ORDER BY catC_livello, catC_id", "ptb_categorieGallery", "catC_id", true)
    CALL CopyTableData(SourceConn, Conn, "ptb_descrittori", "", "des_id", true)
    CALL CopyTableData(SourceConn, Conn, "prel_catGallery_Descrittori", "", "rcd_id", true)
    CALL CopyTableData(SourceConn, Conn, "ptb_gallery", "", "gallery_id", true)
    CALL CopyTableData(SourceConn, Conn, "ptb_immagini", "", "i_id", true)
    CALL CopyTableData(SourceConn, Conn, "prel_descrittori_gallery", "", "rdi_id", true)
    
    SourceConn.close
    Set SourceConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1100
'...........................................................................................
'Importa dati della modulistica nelle nuove categorie delle gallery
'ATTENZIONE: considera gli id dei descrittori dei moduli:
'   1:      Data pubblicazione
'   2:      Link
'   4:      Descrizione
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1100)
if DB.last_update_executed then
	CALL Aggiornamento_1100_import_gallery(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1100_import_gallery(Conn, rs, rst)
	
    dim ReadConn, CatModulistica, sql, Categoria, rsg, rsd, Gallery
	set ReadConn = Server.CreateObject("ADODB.Connection")
    set rsg = Server.CreateObject("ADODB.Recordset")
    set rsd = Server.CreateObject("ADODB.Recordset")

	ReadConn.Open Application("DATA_ConnectionString"), "", ""

    CatModulistica = import_GetCategoriaAlternativa(CatGallery, conn, rs, "MODULISTICA")
    
    'importa categorie moduli
    sql = "SELECT * FROM tb_moduli_cat"
    rs.open sql, ReadConn, adOpenStatic, adLockOptimistic
    while not rs.eof
        Categoria =  import_CATEGORIA(conn, rst, CatGallery, CatModulistica, "MOD_" & rs("cat_mod_id"), rs("cat_mod_nome"), "", "", "", "", rs("cat_mod_id"))
        
        'aggiorna specchio nell'indice
        CALL Index_UpdateItem(conn, CatGallery.tabella, Categoria, false)
        
        'collega i descrittori alla categoria principale anche alla categoria appena inserita
        sql = "DELETE FROM prel_catGallery_descrittori WHERE rcd_categoria_id=" & Categoria
        CALL conn.execute(sql, ,adExecuteNoRecords)
        
        sql = " INSERT INTO prel_catGallery_descrittori (rcd_categoria_id, rcd_descrittore_id, rcd_ordine) " + _
              " SELECT " & Categoria & ", rcd_descrittore_id, rcd_ordine FROM prel_catGallery_descrittori WHERE rcd_categoria_id=" & CatModulistica
        CALL conn.execute(sql, ,adExecuteNoRecords)
        
        'importa dati dei moduli corrispondenti
        sql = "SELECT * FROM tb_moduli WHERE mod_cat=" & rs("cat_mod_id")
        rst.open sql, conn, adOpenStatic, adLockOptimistic
        
        while not rst.eof
            
            'inserisce gallery
            sql = "SELECT * FROM ptb_gallery WHERE gallery_codice='MOD_" & rst("mod_id") & "'"
            rsg.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
            
            if rsg.eof then
                rsg.AddNew
            end if
            rsg("gallery_name_it") = rst("mod_titolo")
            rsg("gallery_idCategoria") = Categoria
            rsg("gallery_codice") = "MOD_" & rst("mod_id")
            rsg("gallery_visibile") = true
            rsg.Update
            Gallery = rsg("gallery_id")
            rsg.close
            
            'inserisce descrittori
            if cString(rst("mod_link"))<>"" then
                CALL Aggiornamento_1100_import_gallery_InsertValoreDescrittore(conn, rsd, 2, Gallery, "rdi_valore_it", rst("mod_link"))
            end if
            if cString(rst("mod_descrizione"))<>"" then
                CALL Aggiornamento_1100_import_gallery_InsertValoreDescrittore(conn, rsd, 4, Gallery, "rdi_memo_it", rst("mod_descrizione"))
            end if
            if cString(rst("mod_data"))<>"" AND IsDate(rst("mod_data")) then
                CALL Aggiornamento_1100_import_gallery_InsertValoreDescrittore(conn, rsd, 1, Gallery, "rdi_valore_it", rst("mod_data"))
            end if
            
            'aggiorna specchio nell'indice
            CALL Index_UpdateItem(conn, "ptb_gallery", Gallery, false)
            rst.movenext
        wend
        
        rst.close
        
        rs.movenext
    wend
    rs.close
    
    ReadConn.close
    set rsg = nothing
    set rsd = nothing
    set ReadConn = nothing
end sub


sub Aggiornamento_1100_import_gallery_InsertValoreDescrittore(conn, rs, des_id, gallery_id, des_field, des_value)
    dim sql
    sql = "SELECT * FROM prel_descrittori_gallery WHERE rdi_gallery_id=" & gallery_id & " AND rdi_descrittore_id=" & des_id
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    if rs.eof then
        rs.AddNew
    end if
    rs("rdi_descrittore_id") = des_id
    rs("rdi_gallery_id") = gallery_id
    rs(des_field) = des_value
    rs.Update
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1101
'...........................................................................................
sql = Aggiornamento__INFO__6(conn)
CALL DB.Execute(sql, 1101)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1101)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1102
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__57(conn)
CALL DB.Execute(sql, 1102)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1103
'...........................................................................................
'aggiorna applicativo gestione contenuti portale in gestione dati associazioni
'...........................................................................................
sql = " UPDATE tb_siti SET " + _
      "     sito_nome = 'Assessorato al turismo [gestione associazioni di categoria] ', " + _
      "     sito_dir = '../Admin/Associazioni', " + _
      "     sito_p1 = 'ASSOCIAZIONI_USER' " + _
      " WHERE id_sito=" & TURISMO_ASSOCIAZIONI
CALL DB.Execute(sql, 1103)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1104
'...........................................................................................
sql = Aggiornamento__INFO__7(conn)
CALL DB.Execute(sql, 1104)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1105
'...........................................................................................
sql = Aggiornamento__INFO__8(conn)
CALL DB.Execute(sql, 1105)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1106
'...........................................................................................
sql = Aggiornamento__INFO__9(conn)
CALL DB.Execute(sql, 1106)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1107
'...........................................................................................
'corregge data di pubblicazione eventi: la porta al giorno della pubblicazione del portale 03/12/2007.
'...........................................................................................
sql = " UPDATE itb_eventi SET eve_pubblData = " & SQL_Date(conn, DateSerial(2007, 12, 3))
CALL DB.Execute(sql, 1107)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1108
'...........................................................................................
sql = Aggiornamento__INFO__10(conn)
CALL DB.Execute(sql, 1108)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1109
'...........................................................................................
sql = Aggiornamento__INFO__11(conn)
CALL DB.Execute(sql, 1109)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1110
'...........................................................................................
sql = Aggiornamento__INFO__12(conn)
CALL DB.Execute(sql, 1110)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1111
'...........................................................................................
sql = Aggiornamento__INFO__13(conn)
CALL DB.Execute(sql, 1111)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1112
'...........................................................................................
sql = Aggiornamento__INFO__14(conn)
CALL DB.Execute(sql, 1112)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1113
'...........................................................................................
'corregge pubblicazione simboli su descrittori
'...........................................................................................
sql = " UPDATE itb_anagrafiche_descrittori SET and_img = '/simboli/' + and_img WHERE IsNull(and_img, '')<>'' AND NOT (and_img LIKE '%simboli/%' )"
CALL DB.Execute(sql, 1113)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1114
'...........................................................................................
'rimuove campo immagine per dotazioni e servizi: non piu' usato, ora vale l'immagine 
'presente nel descrittore eventualmente collegato alla dotazione.
'...........................................................................................
sql = " ALTER TABLE tb_dotazioni DROP COLUMN dotaz_symb; " + _
      " ALTER TABLE tb_servizi DROP COLUMN serv_symb; " + _
      DropObject(conn, "VIEW_Servizi", "VIEW") + _
      " CREATE VIEW dbo.VIEW_Servizi AS " + vbCrLf + _
	  "   SELECT tb_servizi.*, " + vbCrLf + _
	  "          rel_Grp_serv.rel_Grp_serv_id, tb_grp_vis.Grp_Mod_id, rel_Grp_serv.rel_Grp_serv_Ord, " + vbCrLf + _
	  "          (CASE WHEN GETDATE() BETWEEN rel_Grp_serv.serv_valid_from AND ISNULL(rel_Grp_serv.serv_valid_TO, GETDATE()+1) THEN 1 ELSE 0 END) AS VALIDO " + vbCrLf + _
	  "   FROM tb_servizi " + vbCrLf + _
	  "        INNER JOIN rel_Grp_serv ON tb_servizi.serv_id = rel_Grp_serv.rel_Grp_id_serv " + vbCrLf + _
	  "        INNER JOIN tb_grp_vis ON rel_Grp_serv.rel_id_Grp_serv = tb_grp_vis.Grp_id " + vbCrLF + _
      " ; "
CALL DB.Execute(sql, 1114)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1115
'...........................................................................................
'corregge valori dei descrittori booleani rimuovendo quelli impostati a false.
'...........................................................................................
sql = " DELETE FROM irel_anagrafiche_descrTipi WHERE " + _
      "     rad_descrittore_id IN ( SELECT and_id FROM itb_anagrafiche_descrittori WHERE and_tipo = 11 ) AND " + _
      "     (ISNULL(rad_valore_it, '')='' OR IsNull(rad_valore_it, '') LIKE '0' ) "
CALL DB.Execute(sql, 1115)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1116
'...........................................................................................
'corregge stato di blocco dei descrittori "periodo di apertura" delle strutture ricettive
'...........................................................................................
sql = " UPDATE irel_anTipi_descrittori SET rtd_locked=1 WHERE rtd_descrittore_id BETWEEN 250 AND 253 "
CALL DB.Execute(sql, 1116)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1117
'...........................................................................................
sql = Aggiornamento__INFO__15(conn)
CALL DB.Execute(sql, 1117)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1118
'...........................................................................................
'sposta dati da campi bussola a nuovi campi per stampa calendario
'...........................................................................................
sql = " UPDATE itb_eventi SET " + _
      "     eve_descr_calendario_it = eve_bussola_it, " + _
      "     eve_descr_calendario_en = eve_bussola_en, " + _
      "     eve_descr_calendario_fr = eve_bussola_fr, " + _
      "     eve_descr_calendario_de = eve_bussola_de, " + _
      "     eve_descr_calendario_es = eve_bussola_es; " + _
      " ALTER TABLE itb_eventi DROP COLUMN eve_bussola_it, eve_bussola_en, eve_bussola_fr, eve_bussola_de, eve_bussola_es; "
CALL DB.Execute(sql, 1118)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1119
'...........................................................................................
'aggiunge campo su tabella strutture per codifica dell'area come dato informativo
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD nextInfo_area_id INT NULL ; "
CALL DB.Execute(sql, 1119)

if DB.last_update_executed then
	CALL Aggiornamento_1119_import_sito(DB.objConn)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1119_import_sito(Conn)
    dim sql, area
    sql = "SELECT CodAlb FROM tb_loginStru"
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    while not rs.eof
        sql = "SELECT ana_area_id FROM itb_anagrafiche INNER JOIN tb_indirizzario ON itb_anagrafiche.ana_id = tb_indirizzario.idElencoIndirizzi " + _
              " WHERE SyncroTable LIKE 'VIEW_valid_strutture' AND tb_indirizzario.SyncroKey LIKE '" & rs("CodAlb") & "'"
        area = cIntero(GetValueList(conn, rst, sql))
        if area > 0 then
            sql = "UPDATE tb_strutture SET nextInfo_area_id=" & area & " WHERE RegCode LIKE '" & rs("CodAlb") & "'"
            CALL Conn.execute(sql)
        end if
        rs.movenext
    wend
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1120
'...........................................................................................
sql = Aggiornamento__INFO__16(conn)
CALL DB.Execute(sql, 1120)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1120)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1121
'...........................................................................................
'aggiunge relazione tra aree e strutture ricettive
'...........................................................................................
sql = SQL_AddForeignKey(conn, "tb_Strutture", "nextInfo_area_id", "itb_aree", "are_id", false, "")
CALL DB.Execute(sql, 1121)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1122
'...........................................................................................
'aggiunge campo su tabella associazioni per codifica dell'area come dato informativo
'...........................................................................................
sql = " ALTER TABLE tb_assoc ADD asc_nextInfo_area_id INT NULL ; "
CALL DB.Execute(sql, 1122)

if DB.last_update_executed then
	CALL Aggiornamento_1122_import_sito(DB.objConn)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1122_import_sito(Conn)
    dim sql, area
    sql = "SELECT * FROM tb_assoc"
    rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    while not rs.eof
        
        sql = "SELECT ana_area_id FROM itb_anagrafiche INNER JOIN tb_indirizzario ON itb_anagrafiche.ana_id = tb_indirizzario.idElencoIndirizzi " + _
              " WHERE SyncroTable LIKE 'tb_assoc' AND tb_indirizzario.SyncroKey LIKE '" & rs("asc_id") & "'"
        rs("asc_nextInfo_area_id") = cIntero(GetValueList(conn, rst, sql))
        rs.update
        rs.movenext
    wend
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1123
'...........................................................................................
'rimuove relazione tra aree e strutture ricettive
'...........................................................................................
sql = SQL_RemoveForeignKey(conn, "irel_aree_comuni", "", "", true, "FK_irel_aree_comuni_itb_aree") + _
      SQL_RemoveForeignKey(conn, "irel_aree_comuni", "", "", true, "FK_irel_aree_comuni_tb_localita") + _
      SQL_RemoveForeignKey(conn, "irel_aree_localita", "", "", true, "FK_irel_aree_localita_itb_aree") + _
      SQL_RemoveForeignKey(conn, "irel_aree_localita", "", "", true, "FK_irel_aree_localita_tb_localita") + _
      DropObject(conn, "irel_aree_comuni", "TABLE") + _
      DropObject(conn, "irel_aree_localita", "TABLE")
CALL DB.Execute(sql, 1123)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1124
'...........................................................................................
sql = Aggiornamento__INFO__17(conn)
CALL DB.Execute(sql, 1124)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1125
'...........................................................................................
'corregge impostazioni applicativi passport per gestione permessi aggiuntivi
'...........................................................................................
sql = " UPDATE tb_siti SET " + _
            " sito_prmEsterni_admin = '../../Admin/Passport/PassportAdmin.asp', " + _
            " sito_prmEsterni_sito = '../../Admin/Passport/PassportSito.asp' " + _
        " WHERE id_sito IN (SELECT mod_applicazione_id FROM tb_modelli)"
CALL DB.Execute(sql, 1125)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1126
'...........................................................................................
'corregge impostazioni applicativi passport per gestione permessi aggiuntivi
'...........................................................................................
sql = " DELETE FROM tb_siti WHERE id_sito=" & TURISMO_PASSPORT
CALL DB.Execute(sql, 1126)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1127
'...........................................................................................
'rimuove vecchie tabelle non piu' utilizzate di vecchi contenuti
'...........................................................................................
sql = SQL_RemoveForeignKey(conn, "tb_promozioni", "", "", true, "FK_tb_Promozioni_tb_cat_promo") + _
      SQL_RemoveForeignKey(conn, "tb_informazioni", "", "", true, "FK_tb_informazioni_tb_info_cat") + _
      SQL_RemoveForeignKey(conn, "tb_moduli", "", "", true, "FK_tb_moduli_tb_moduli_cat") + _
      DropObject(conn, "tb_cat_promo", "TABLE") + _
      DropObject(conn, "tb_promozioni", "TABLE") + _
      DropObject(conn, "tb_info_cat", "TABLE") + _
      DropObject(conn, "tb_informazioni", "TABLE") + _
      DropObject(conn, "tb_moduli", "TABLE") + _
      DropObject(conn, "tb_moduli_cat", "TABLE") + _
      DropObject(conn, "tb_turismo_news", "TABLE")
CALL DB.Execute(sql, 1127)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1128
'...........................................................................................
'aggiunge applicativo per export dati NEXT-info (ex bussola)
'...........................................................................................
sql = " INSERT INTO tb_siti(id_sito, sito_nome, sito_amministrazione, sito_dir, sito_p1) " + _
      " VALUES (" & NEXTINFO_EXPORT & ", 'NEXT-info export [export dati informativi]', 1, '../NextInfo_Export', 'EXPORT_USER') ; " + _
	  " INSERT INTO rel_admin_sito (sito_id, admin_id, rel_as_permesso) " + _
	  " SELECT DISTINCT " & NEXTINFO_EXPORT & ", admin_id, 1 FROM rel_admin_sito WHERE sito_id = " & NEXTINFO & " AND " + _
	  " admin_id IN (SELECT adm_id FROM irel_admin WHERE adm_permesso >= 3 OR adm_area_id = 238) " '238 = Area Apt di Venezia
CALL DB.Execute(sql, 1128)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1129
'...........................................................................................
'corregge import dati delle spiagge per importare testi in francese, tedesco e spagnolo.
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1129)
if DB.last_update_executed then
    CALL update_import_SPIAGGE(CODE_APTBC)
	CALL update_import_SPIAGGE(CODE_APTJE)
    CALL update_import_SPIAGGE(CODE_APTVE)
	CALL update_import_SPIAGGE(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1130
'...........................................................................................
'corregge import dati dei luoghi per importare testi in francese, tedesco e spagnolo.
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1130)
if DB.last_update_executed then
    CALL update_import_LUOGHI(CODE_APTBC)
	CALL update_import_LUOGHI(CODE_APTJE)
    CALL update_import_LUOGHI(CODE_APTVE)
	CALL update_import_LUOGHI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1131
'...........................................................................................
'corregge import dati degli eventi per importare testi in francese, tedesco e spagnolo.
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1131)
if DB.last_update_executed then
    CALL update_import_EVENTI(CODE_APTBC)
	CALL update_import_EVENTI(CODE_APTJE)
    CALL update_import_EVENTI(CODE_APTVE)
	CALL update_import_EVENTI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1132
'...........................................................................................
'corregge import dati delle notizie utili per importare testi in francese, tedesco e spagnolo.
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1132)
if DB.last_update_executed then
    CALL update_import_NOTIZIE_UTILI(CODE_APTBC)
	CALL update_import_NOTIZIE_UTILI(CODE_APTJE)
    CALL update_import_NOTIZIE_UTILI(CODE_APTVE)
	CALL update_import_NOTIZIE_UTILI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1133
'...........................................................................................
'corregge import dati dei locali e servizi per importare testi in francese, tedesco e spagnolo.
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1133)
if DB.last_update_executed then
    CALL update_import_LOCALI_SERVIZI(CODE_APTBC)
	CALL update_import_LOCALI_SERVIZI(CODE_APTJE)
    CALL update_import_LOCALI_SERVIZI(CODE_APTVE)
	CALL update_import_LOCALI_SERVIZI(CODE_APTCH)
end if
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1134
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__58(conn)
CALL DB.Execute(sql, 1134)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1135
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__59(conn)
CALL DB.Execute(sql, 1135)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1136
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__60(conn)
CALL DB.Execute(sql, 1136)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1137
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__61(conn)
CALL DB.Execute(sql, 1137)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1137)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1138
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__62(conn)
CALL DB.Execute(sql, 1138)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1139
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__63(conn)
CALL DB.Execute(sql, 1139)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1140
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__64(conn)
CALL DB.Execute(sql, 1140)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1141
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__65(conn)
CALL DB.Execute(sql, 1141)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1142
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__66(conn)
CALL DB.Execute(sql, 1142)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1143
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__67(conn)
CALL DB.Execute(sql, 1143)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1144
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__68(conn)
CALL DB.Execute(sql, 1144)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1145
'...........................................................................................
sql = Aggiornamento__INFO__18(conn)
CALL DB.Execute(sql, 1145)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1146
'...........................................................................................
sql = Aggiornamento__INFO__19(conn)
CALL DB.Execute(sql, 1146)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1147
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__69(conn)
CALL DB.Execute(sql, 1147)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1148
'...........................................................................................
'	corregge valore colonna storico per la gestione del "progressivo"
'...........................................................................................
sql = " UPDATE tb_str_logs SET str_log_codAlb = RTRIM(LTRIM(str_log_codAlb)), str_log_progressivo=NULL; "
CALL DB.Execute(sql, 1148)
if DB.last_update_executed then
	CALL Aggiornamento_1148_GenerazioneProgressivo(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1148_GenerazioneProgressivo(conn, rs)
    dim CurrentRegCode, Progressivo
    sql = " SELECT * FROM tb_str_logs " + _
		  " WHERE Str_log_des LIKE '%cancellazione completa struttura%' OR " + _
		  		" Str_log_des LIKE '%registrazione validata%' OR " + _
				" (str_log_id IN (SELECT MAX(str_log_id) FROM tb_str_logs WHERE (Str_log_des LIKE '%Modifica%' AND Str_log_des NOT LIKE '%compilazione%') " + _
				"	AND str_log_data <= CONVERT(DATETIME, '2008-06-20 00:00:00', 102) GROUP BY str_log_codalb, CONVERT(nvarchar(30), str_log_data, 102) ) ) " + _
		  " ORDER BY str_log_codalb, str_log_id "
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    CurrentRegCode = ""
    Progressivo = 0
    while not rs.eof
        if CurrentRegCode <> Trim(rs("str_log_codAlb")) then
            CurrentRegCode = Trim(rs("str_log_codAlb"))
            Progressivo = 0
        else
            Progressivo = Progressivo + 1
        end if
        rs("str_log_progressivo") = Progressivo
        rs.Update
        
        rs.movenext
    wend
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1149
'...........................................................................................
'aggiunge campo con modello su log delle strutture
'...........................................................................................
sql = " ALTER TABLE tb_str_logs ADD str_log_modello INT NULL; " + _
      " UPDATE tb_str_logs SET str_log_modello = (SELECT TOP 1 modello FROM tb_loginstru WHERE CodAlb LIKE LEFT(tb_str_logs.str_log_codalb, 3) + '%' ) ; "
CALL DB.Execute(sql, 1149)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1150
'...........................................................................................
'	corregge valore colonna storico per la gestione del "progressivo" causa data errata in precedente update
'...........................................................................................
sql = " UPDATE tb_str_logs SET str_log_codAlb = RTRIM(LTRIM(str_log_codAlb)), str_log_progressivo=NULL; "
CALL DB.Execute(sql, 1150)
if DB.last_update_executed then
	CALL Aggiornamento_1150_GenerazioneProgressivo(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1150_GenerazioneProgressivo(conn, rs)
    dim CurrentRegCode, Progressivo
    sql = " SELECT * FROM tb_str_logs " + _
		  " WHERE Str_log_des LIKE '%cancellazione completa struttura%' OR " + _
		  		" Str_log_des LIKE '%registrazione validata%' OR " + _
				" (str_log_id IN (SELECT MAX(str_log_id) FROM tb_str_logs WHERE (Str_log_des LIKE '%Modifica%' OR Str_log_des LIKE '%Inserimento nuova struttura%' AND Str_log_des NOT LIKE '%compilazione%') " + _
				"	AND str_log_data <= CONVERT(DATETIME, '2007-06-20 00:00:00', 102) GROUP BY str_log_codalb, CONVERT(nvarchar(30), str_log_data, 102) ) ) " + _
		  " ORDER BY str_log_codalb, str_log_id "
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
    CurrentRegCode = ""
    Progressivo = 1
    while not rs.eof
        if CurrentRegCode <> Trim(rs("str_log_codAlb")) then
            CurrentRegCode = Trim(rs("str_log_codAlb"))
            Progressivo = 1
        else
            Progressivo = Progressivo + 1
        end if
        rs("str_log_progressivo") = Progressivo
        rs.Update
        
        rs.movenext
    wend
    rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1151
'...........................................................................................
'	corregge trigger alla struttura di log per il calcolo del progressivo
'...........................................................................................
sql = DropObject(conn, "tb_str_logs_INSERT", "TRIGGER") + _
	  " CREATE TRIGGER dbo.tb_str_logs_INSERT ON tb_str_logs AFTER INSERT AS " + vbCrLf + _
      "     DECLARE @REGCODE nvarchar(13) " + vbCrLF + _
      "     DECLARE @LAST INT " + vbCrLF + _
      "     DECLARE @PROGRESSIVO INT " + vbCrLF + _
      vbCrLf + _
      "     SELECT @REGCODE = RTRIM(LTRIM(str_log_CodAlb)), @LAST=str_log_id " + vbCrLF + _
      "         FROM INSERTED WHERE ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione completa struttura%')) " + vbCrLf + _
      vbCrLF + _
      "     if (@REGCODE <> '') " + vbCRLF + _
      "         BEGIN " + vbCrLF + _
      "             IF (EXISTS(SELECT * FROM tb_str_logs WHERE RTRIM(LTRIM(str_log_CodAlb)) LIKE @REGCODE AND str_log_id <> @LAST AND IsNull(str_log_progressivo,0)<>0 )) " + vbCrLf + _
      "                 BEGIN " + vbCrLF + _
      "                     SELECT @PROGRESSIVO = MAX(str_log_progressivo) FROM tb_str_logs " + vbCrLF + _
      "                         WHERE RTRIM(LTRIM(str_log_CodAlb)) LIKE @REGCODE AND str_log_id <> @LAST " + vbCrLf + _
      "                     SET @PROGRESSIVO = @PROGRESSIVO + 1 " + vbCrLF + _
      "                 END " + vbCrLf + _
      "             ELSE " + vbCrLF + _
      "                 BEGIN " + vbCrLF + _
      "                     SET @PROGRESSIVO = 1 " + vbCrLF + _
      "                 END " + vbCrLF + _
      "             UPDATE tb_str_logs SET str_log_progressivo=@PROGRESSIVO WHERE str_log_id=@LAST " + vbCrLF + _
      "         END " + vbCrLF + _
      " ; "
CALL DB.Execute(sql, 1151)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1152
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__70(conn)
CALL DB.Execute(sql, 1152)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1153
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__71(conn)
CALL DB.Execute(sql, 1153)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1154
'...........................................................................................
'	corregge area non impostata su strutture ricettive pubblicate
'...........................................................................................
sql = " UPDATE tb_strutture SET nextinfo_area_id = ( SELECT TOP 1 nextinfo_area_id FROM VIEW_records_strutture " + _
												   " WHERE VIEW_records_strutture.regcode = tb_strutture.regcode AND ISNULL(nextinfo_area_id, 164)<>164 " + _
												   " ORDER BY str_id DESC ) " + _
      " WHERE ISNULL(nextinfo_area_id, 164)=164 AND EXISTS ( SELECT CodAlb FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginstru.modello = tb_modelli.mod_id " + _
	  													   " WHERE tb_loginstru.codalb = tb_strutture.regcode AND ISNULL(tb_modelli.mod_portale_pubblica,0)=1 ) ; " + _
	  " UPDATE tb_strutture SET nextinfo_area_id = (SELECT ana_area_id FROM itb_anagrafiche WHERE ana_codice LIKE tb_strutture.regcode) " + _
	  "	WHERE ISNULL(nextinfo_area_id, 164)=164 AND EXISTS ( SELECT CodAlb FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginstru.modello = tb_modelli.mod_id " + _
	  													   " WHERE tb_loginstru.codalb = tb_strutture.regcode AND ISNULL(tb_modelli.mod_portale_pubblica,0)=1 ) ; "
CALL DB.Execute(sql, 1154)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1155
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__72(DB, 1155)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1156
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__73(conn)
CALL DB.Execute(sql, 1156)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1157
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__74(conn)
CALL DB.Execute(sql, 1157)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1158
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__75(conn)
CALL DB.Execute(sql, 1158)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1158)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1159
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__76(conn)
CALL DB.Execute(sql, 1159)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1160
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__77(conn)
CALL DB.Execute(sql, 1160)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1161
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__78(conn)
CALL DB.Execute(sql, 1161)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1162
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__79(conn)
CALL DB.Execute(sql, 1162)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1163
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__80(conn)
CALL DB.Execute(sql, 1163)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1164
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__81(conn)
CALL DB.Execute(sql, 1164)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1164)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1165
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__82(conn)
CALL DB.Execute(sql, 1165)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1166
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__83(conn)
CALL DB.Execute(sql, 1166)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1167
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__84(conn)
CALL DB.Execute(sql, 1167)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1168
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__85(conn)
CALL DB.Execute(sql, 1168)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1169
'...........................................................................................
sql = Aggiornamento__INFO__20(conn)
CALL DB.Execute(sql, 1169)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1170
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__86(conn)
CALL DB.Execute(sql, 1170)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1171
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__87(DB, 1171)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1172
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__88(conn)
CALL DB.Execute(sql, 1172)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1173
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__89(conn)
CALL DB.Execute(sql, 1173)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1173)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1174
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__90(DB, rs, 1174)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1175
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__91(conn)
CALL DB.Execute(sql, 1175)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1175)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1176
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__92(conn)
CALL DB.Execute(sql, 1176)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1177
'...........................................................................................
'	crea colonne per revisione codice regionale
'...........................................................................................
sql = " ALTER TABLE tb_loginStru ADD " + _
	  " 	CodAlb_Old nvarchar(12), " + _
	  "		CodAlb_Regione nvarchar(12) " + _
	  " ; " + _
	  " UPDATE tb_loginStru SET CodAlb_Old = CodAlb, CodAlb_regione = dbo.fn_regcode_for_regione(CodAlb, modello) ; "
CALL DB.Execute(sql, 1177)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1178
'...........................................................................................
'	corregge dimensione codice proprietario e codice tipologia
'...........................................................................................
sql = " ALTER TABLE tb_strutture ALTER COLUMN Cod_proprietario nvarchar(12); " + _
	  " ALTER TABLE tb_strutture ALTER COLUMN Cod_tipologia nvarchar(12); "
CALL DB.Execute(sql, 1178)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1178)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1179
'...........................................................................................
'	modifica stored procedure per sostituzione codice regionale
'...........................................................................................
sql = DropObject(conn, "spstr_REPLACE_REGCODE", "PROCEDURE") + _
	  " CREATE PROCEDURE dbo.spstr_REPLACE_REGCODE(" + vbCrLF + _
	  "    @REGCODE_OLD nvarchar(12), " + vbCrLF + _
	  "    @REGCODE_NEW nvarchar(12) " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "    BEGIN TRAN " + vbCrLF + _
	  "        SELECT @REGCODE_OLD, @REGCODE_NEW " + vbCrLF + _
	  "        --aggiornamento gestione strutture ricettive " + vbCrLf +_
	  "        UPDATE tb_loginstru SET CodAlb = @REGCODE_NEW WHERE CodAlb = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE rel_assoc_stru SET CodAlb_rel = @REGCODE_NEW WHERE CodAlb_rel = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE tb_str_logs SET str_log_CodAlb = @REGCODE_NEW WHERE str_log_codAlb = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE tb_stru_gest SET RegCode = @REGCODE_NEW WHERE RegCode = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE tb_strutture SET RegCode = @REGCODE_NEW WHERE RegCode = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE tb_strutture SET Cod_Tipologia = @REGCODE_NEW WHERE Cod_Tipologia = @REGCODE_OLD " + vbCrLF + _
	  "        UPDATE tb_strutture SET Cod_proprietario = @REGCODE_NEW WHERE Cod_proprietario = @REGCODE_OLD " + vbCrLF + _
	  "        --aggiornamento strutture sincronizzate con NEXT-INFO " + vbCrLF + _
	  "        UPDATE itb_anagrafiche SET ana_codice = @REGCODE_NEW WHERE ana_codice LIKE @REGCODE_OLD " + vbCrLF + _
	  "        --aggiornamento contatti sincronizzati con NEXT-COM " + vbCrLF + _
	  "        UPDATE tb_indirizzario SET SyncroKey = @REGCODE_NEW WHERE SyncroKey LIKE @REGCODE_OLD " + vbCrLF + _
	  "    COMMIT TRAN "
CALL DB.Execute(sql, 1179)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1180
'...........................................................................................
'	modifica stored procedure per nuova generazione codice regionale
'...........................................................................................
sql = 	DropObject(conn, "spstr_INSERT_NEW", "PROCEDURE") + _
		" CREATE PROCEDURE dbo.spstr_INSERT_NEW( " + vbCrLF + _
        "     @DENOMINAZIONE nvarchar(60), " + vbCrLF + _
        "     @COMUNE nvarchar(6), " + vbCrLF + _
        "     @MODELLO int, " + vbCrLF + _
        "     @TIPO int, " + vbCrLF + _
        "     @UTENTE nvarchar(50), " + vbCrLF + _
        "     @STR_ID int OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
        "     DECLARE @DATAMOD smalldatetime " + vbCrLF + _
        "     SET @DATAMOD = GETDATE() " + vbCrLF + _
        "     SET @DATAMOD = CONVERT(DATETIME, str(YEAR(@DATAMOD)) + '-' + str(MONTH(@DATAMOD)) + '-' + str(DAY(@DATAMOD)) + ' 00:00:00', 102) " + vbCrLF + _
        vbCrLF + _
        "     --calcolo codice regionale " + vbCrLF + _
        "     DECLARE @FIRST_LETTER nvarchar(3) " + vbCrLF + _
        "     DECLARE @VAR_CODEPART int " + vbCrLF + _
        "     DECLARE @VAR_CODEPART_LENGTH int " + vbCrLF + _
        vbCrLF + _
        "     --recupera radicie del codice regionale " + vbCrLF + _
        "     SELECT @FIRST_LETTER=Mod_FirstLT_Regcode FROM tb_Modelli WHERE Mod_ID= @MODELLO " + vbCrLF + _
        vbCrLF + _
        "     --compone prima parte del codice regionale " + vbCrLF + _
        "     SET @REGCODE = @FIRST_LETTER + @COMUNE " + vbCrLF + _
        vbCrLF + _
        "     --calcola lunghezza parte incrementale del codice " + vbCrLF + _
        "     IF (LEN(@FIRST_LETTER)) > 2 " + vbCrLF + _
        "         SET @VAR_CODEPART_LENGTH = 12 - LEN(@REGCODE) " + vbCrLF + _
        "     ELSE " + vbCrLF + _
        "         SET @VAR_CODEPART_LENGTH = 11 - LEN(@REGCODE) " + vbCrLF + _
        vbCrLF + _
        "     --recupera parte incrementale del codice (ultimo inserito) " + vbCrLF + _
        "     SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(RTRIM(CodAlb), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
        "         FROM tb_loginstru " + vbCrLF + _
        "         WHERE Modello = @MODELLO AND CodAlb LIKE (@REGCODE + '%') " + vbCrLF + _
        vbCrLF + _
        "     --incrementa parte variabile " + vbCrLF + _
        "     SET @VAR_CODEPART = @VAR_CODEPART + 1 " + vbCrLF + _
        vbCrLF + _
        "     --compone codice regionale definitivo " + vbCrLF + _
        "     SET @REGCODE = @REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART AS NVARCHAR(12)) " + vbCrLF + _
        vbCrLF + _
        "     --inserimento su tabella tb_loginStru " + vbCrLF + _
        "     INSERT INTO tb_loginstru (CODALB  , Modello , struttura_attiva, custom_dichiarazione ) " + vbCrLF + _
        "         VALUES(               @REGCODE, @MODELLO, 0               , 0                    ) " + vbCrLF + _
        vbCrLF + _
        "     --inserimento record su tabella strutture " + vbCrLF + _
        "     INSERT INTO tb_Strutture (Denominazione , RegCode , DataModifica, UtenteModifica, Tipo , Comune , prezziEuro, record_validato, avviso_inviato, online_dic_presentata, online_dic_completata, online_dic_annullata, online_dichiarazione_id ) " + vbCrLF + _
        "         VALUES(               @DENOMINAZIONE, @REGCODE, @DATAMOD    , @UTENTE       , @TIPO, @COMUNE, 1         , 0              , 0             , 0                    , 0                    , 0                   , NULL                    ) " + vbCrLF + _
        vbCrLF + _
        "     --legge STR_ID " + vbCrLF + _
        "     SELECT @STR_ID = MAX(Str_ID) FROM tb_Strutture WHERE RegCode = @REGCODE " + vbCrLF + _
        vbCrLF + _
        "     --inserimento record su tabella stru_gest " + vbCrLF + _
        "     INSERT INTO tb_Stru_Gest (str_ID , RegCode , DataModifica, immobile_loc, azienda_loc, F_CH_TMP, F_REVOCA_CL, F_REVOCA_LIC) " + vbCrLF + _
        "         VALUES(               @STR_ID, @REGCODE, @DATAMOD    , 0           , 0          , 'N'     , 'N'        , 'N'         ) " + vbCrLF + _
        vbCrLF + _
        "     --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato)" + vbCrLF + _
        "     EXEC spstr_UPDATE_tb_loginstru @RegCode " + vbCrLF + _
        " ; "
CALL DB.Execute(sql, 1180)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1181
'...........................................................................................
'	corregge dimensione codice proprietario e codice tipologia
'...........................................................................................
sql = DropObject(conn, "fn_regcode_for_regione", "FUNCTION") + _
      DropObject(conn, "fn_regcode_from_regione", "FUNCTION")
CALL DB.Execute(sql, 1181)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1182
'...........................................................................................
'	crea nuova stored procedure per il calcolo del codice regionale
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_NEW_REGCODE(" + vbCrLF + _
	  "		@COMUNE nvarchar(6), " + vbCrLF + _
      "		@MODELLO int " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " RETURNS nvarchar(12) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " BEGIN " + vbCrLF + _
	  "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
	  "     DECLARE @FIRST_LETTER nvarchar(3) " + vbCrLF + _
	  "     DECLARE @VAR_CODEPART int " + vbCrLF + _
	  "     DECLARE @CODE_LENGHT int " + vbCrLF + _
      "     DECLARE @VAR_CODEPART_LENGTH int " + vbCrLF + _
      vbCrLF + _
      "     --recupera radicie del codice regionale " + vbCrLF + _
	  "     SELECT @FIRST_LETTER = Mod_FirstLT_Regcode FROM tb_Modelli WHERE Mod_ID= @MODELLO " + vbCrLF + _
      vbCrLF + _
      "     --compone prima parte del codice regionale " + vbCrLF + _
      "     SET @REGCODE = @FIRST_LETTER + @COMUNE " + vbCrLF + _
      vbCrLF + _
      "     --calcola lunghezza parte incrementale del codice " + vbCrLF + _
      "     IF (LEN(@FIRST_LETTER)) > 1 " + vbCrLF + _
      "         SET @CODE_LENGHT = 12 " + vbCrLF + _
      "     ELSE " + vbCrLF + _
	  "         SET @CODE_LENGHT = 11 " + vbCrLF + _
	  "     SET @VAR_CODEPART_LENGTH = @CODE_LENGHT - LEN(@REGCODE) " + vbCrLF + _
      vbCrLF + _
      "     --recupera parte incrementale del codice (ultimo inserito) " + vbCrLF + _
      "     SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(RTRIM(CodAlb), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
      "         FROM tb_loginstru " + vbCrLF + _
      "         WHERE CodAlb LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "               LEN(CodAlb) = @CODE_LENGHT " + vbCrLF + _
      vbCrLF + _
      "     --incrementa parte variabile " + vbCrLF + _
      "     SET @VAR_CODEPART = @VAR_CODEPART + 1 " + vbCrLF + _
      vbCrLF + _
      "     --compone codice regionale definitivo " + vbCrLF + _
      "     SET @REGCODE = @REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART AS NVARCHAR(12)) " + vbCrLF + _
      vbCrLF + _
	  "     RETURN @REGCODE " + vbCrLF + _
	  " END " + vbCrLF
CALL DB.Execute(sql, 1182)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1183
'...........................................................................................
'	modifica stored procedure per nuova generazione codice regionale
'...........................................................................................
sql = 	DropObject(conn, "spstr_INSERT_NEW", "PROCEDURE") + _
		" CREATE PROCEDURE dbo.spstr_INSERT_NEW( " + vbCrLF + _
        "     @DENOMINAZIONE nvarchar(60), " + vbCrLF + _
        "     @COMUNE nvarchar(6), " + vbCrLF + _
        "     @MODELLO int, " + vbCrLF + _
        "     @TIPO int, " + vbCrLF + _
        "     @UTENTE nvarchar(50), " + vbCrLF + _
        "     @STR_ID int OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
        "     DECLARE @DATAMOD smalldatetime " + vbCrLF + _
        "     SET @DATAMOD = GETDATE() " + vbCrLF + _
        "     SET @DATAMOD = CONVERT(DATETIME, str(YEAR(@DATAMOD)) + '-' + str(MONTH(@DATAMOD)) + '-' + str(DAY(@DATAMOD)) + ' 00:00:00', 102) " + vbCrLF + _
        vbCrLF + _
        "     --calcolo codice regionale " + vbCrLF + _
        "     SELECT @REGCODE = dbo.fn_NEW_REGCODE(@COMUNE, @MODELLO) " + vbCrLf + _
        vbCrLF + _
        "     --inserimento su tabella tb_loginStru " + vbCrLF + _
        "     INSERT INTO tb_loginstru (CODALB  , Modello , struttura_attiva, custom_dichiarazione ) " + vbCrLF + _
        "         VALUES(               @REGCODE, @MODELLO, 0               , 0                    ) " + vbCrLF + _
        vbCrLF + _
        "     --inserimento record su tabella strutture " + vbCrLF + _
        "     INSERT INTO tb_Strutture (Denominazione , RegCode , DataModifica, UtenteModifica, Tipo , Comune , prezziEuro, record_validato, avviso_inviato, online_dic_presentata, online_dic_completata, online_dic_annullata, online_dichiarazione_id ) " + vbCrLF + _
        "         VALUES(               @DENOMINAZIONE, @REGCODE, @DATAMOD    , @UTENTE       , @TIPO, @COMUNE, 1         , 0              , 0             , 0                    , 0                    , 0                   , NULL                    ) " + vbCrLF + _
        vbCrLF + _
        "     --legge STR_ID " + vbCrLF + _
        "     SELECT @STR_ID = MAX(Str_ID) FROM tb_Strutture WHERE RegCode = @REGCODE " + vbCrLF + _
        vbCrLF + _
        "     --inserimento record su tabella stru_gest " + vbCrLF + _
        "     INSERT INTO tb_Stru_Gest (str_ID , RegCode , DataModifica, immobile_loc, azienda_loc, F_CH_TMP, F_REVOCA_CL, F_REVOCA_LIC) " + vbCrLF + _
        "         VALUES(               @STR_ID, @REGCODE, @DATAMOD    , 0           , 0          , 'N'     , 'N'        , 'N'         ) " + vbCrLF + _
        vbCrLF + _
        "     --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato)" + vbCrLF + _
        "     EXEC spstr_UPDATE_tb_loginstru @RegCode " + vbCrLF + _
        " ; "
CALL DB.Execute(sql, 1183)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1184
'...........................................................................................
'	modifica dati per costruzione codice regionale
'...........................................................................................
sql = " UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'CP' WHERE mod_id = 22; " + _
	  " UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'MP' WHERE mod_id = 30; " + _
	  " UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'IP' WHERE mod_id = 34; " + _
	  " UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'T' WHERE mod_id = 35; "
CALL DB.Execute(sql, 1184)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1185
'...........................................................................................
'	rimuove relazione tra modelli ed uffici apt
'...........................................................................................
sql = " ALTER TABLE rel_mod_apt_uffici DROP CONSTRAINT FK_rel_mod_apt_uffici_tb_modelli ; " + _
	  " ALTER TABLE rel_mod_apt_uffici DROP CONSTRAINT FK_rel_mod_apt_uffici_tb_apt_uffici ; " + _
	  DropObject(conn, "rel_mod_apt_uffici", "TABLE")
CALL DB.Execute(sql, 1185)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1185)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1186
'...........................................................................................
'	genera nuovi codici regionali
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1186)
if DB.last_update_executed then
	CALL Aggiornamento_1186_RigenerazioneCodiciRegionali(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1186_RigenerazioneCodiciRegionali(conn, rs)
	
	sql = "SELECT RegCode, Modello, comune FROM VIEW_testata_strutture WHERE modello IN (22, 23, 24, 25, 26, 30, 34, 35, 37, 38, 39) ORDER BY CodAlb_Old "
	rs.open sql, conn, adOpenDynamic, adLockOptimistic
	sql = ""
	while not rs.eof
		
		'calcola nuovo codice regionale e imposta nuovo codice regionale
		sql = sql + _
			  " DECLARE @NEW_CODE nvarchar(12) " + vbCrLF + _
			  " SELECT @NEW_CODE = dbo.fn_NEW_REGCODE('" & rs("comune") & "', '" & rs("modello") & "') " + vbCrLF + _
			  " EXEC spstr_REPLACE_REGCODE '" & Trim(rs("RegCode")) & "', @NEW_CODE " + vbCrLF + _
			  " ; " + vbCrLF
		
		rs.movenext
	wend
	rs.close
	
	'esegue aggiornamento
	CALL ExecuteMultipleSql(conn, sql, true)
	
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1187
'...........................................................................................
'	modifica dati per costruzione codice regionale
'...........................................................................................
sql = " UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'M' WHERE mod_id = 36; "
CALL DB.Execute(sql, 1187)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1188
'...........................................................................................
'	modifica files e nomi dei files per allineamento con codici regionali
'...........................................................................................
sql = " SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1188)
if DB.last_update_executed then
	CALL Aggiornamento_1188_AllineamentoArchivioCodiciRegionali(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1188_AllineamentoArchivioCodiciRegionali(conn, rs)
	'ridenominazione dei files.
	dim fso, DirModelli, FileModello, RegCode, oldName
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set DirModelli = fso.GetFolder(Application("IMAGE_PATH") & "dichiarazioni_strutture\")
	
	for each FileModello in DirModelli.Files
		oldName = FileModello.Name
		RegCode = Trim(left(FileModello.Name, instr(1, FileModello.Name, "_", vbTextCompare) - 1)) %>
		<!-- 
			<%= FileModello.Name %>
			<%= RegCode %>
			<% sql = "SELECT CodAlb FROM tb_loginStru WHERE LTRIM(RTRIM(CodAlb_Old)) LIKE '" & RegCode & "' AND NOT ( LTRIM(RTRIM(CodAlb)) LIKE LTRIM(RTRIM(CodAlb_Old)) ) "
			rs.open sql, conn, adOpenDynamic, adLockOptimistic
			if not rs.eof then
				FileModello.Name = replace(FileModello.Name, RegCode, rs("CodAlb")) 
				if lcase(oldName)<>lcase(FileModello.name) then%>
					<!-- 
					<%= FileModello.Name %>
					<%= rs("CodAlb") %>
					<% sql = " UPDATE tb_strutture SET archivio_modello_dichiarazione= '" & FileModello.Name & "' " + _
							 " WHERE archivio_modello_dichiarazione LIKE '" & oldName & "' " + vbCrLF + _
							 " UPDATE tb_strutture SET archivio_tabella_prezzi= '" & FileModello.Name & "' " + _
							 " WHERE archivio_tabella_prezzi LIKE '" & oldName & "' "
					CALL conn.execute(sql)
				end if
			end if %>
		-->
		<% rs.close
	next
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1189
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__92(conn)
CALL DB.Execute(sql, 1189)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1190
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__95(conn)
CALL DB.Execute(sql, 1190)
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
'AGGIORNAMENTO 1191
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__96(conn)
CALL DB.Execute(sql, 1191)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1192
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__97(conn)
CALL DB.Execute(sql, 1192)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1193
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__98(conn)
CALL DB.Execute(sql, 1193)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1194
'...........................................................................................
'	aggiorna funzione per il calcolo del totale bagni
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_bagni", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_calcola_totale_bagni (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      "  " + vbCrLF + _
      " AS  " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         /* " + vbCrLF + _
      "         Totale generale bagni:				207 " + vbCrLF + _
      "         Camerini bagno chiusi:				317 " + vbCrLF + _
      "         Servizi igienici per singoli equipaggi:		326 " + vbCrLF + _
      "         Servizi igienici per disabili:			318 " + vbCrLF + _
      "         Unit abitative con servizi igienici:		303 " + vbCrLF + _
      "         */ " + vbCrLF + _
      vbCrLF + _
      "         DECLARE @TOT_BAGNI int " + vbCrLF + _
      "         SELECT @TOT_BAGNI = ISNULL(SUM(rel_str_dotaz_valore), 0)" + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND  rel_grp_id_dotaz IN (207, 317, 326, 318) " + vbCrLF + _
      vbCrLF + _
      "         SELECT @TOT_BAGNI = @TOT_BAGNI + ISNULL(SUM(rel_str_dotaz_valore), 0) " + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val IN (1, 3)  AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND rel_grp_id_dotaz IN (303)  " + vbCrLF + _
      "         RETURN @TOT_BAGNI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 1194)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1195
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__99(conn)
CALL DB.Execute(sql, 1195)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1196
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__100(conn)
CALL DB.Execute(sql, 1196)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1196)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1197
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__101(conn)
CALL DB.Execute(sql, 1197)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1198
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__102(conn)
CALL DB.Execute(sql, 1198)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1199
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__103(conn)
CALL DB.Execute(sql, 1199)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1199)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1200
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__104(conn)
CALL DB.Execute(sql, 1200)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1201
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__105(conn)
CALL DB.Execute(sql, 1201)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.RebuildIndex_RefreshContents("tb_webs", "id_webs")
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1202
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__106(conn)
CALL DB.Execute(sql, 1202)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1203
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__107(conn)
CALL DB.Execute(sql, 1203)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1204
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__108(conn)
CALL DB.Execute(sql, 1204)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1205
'...........................................................................................
'	corregge funzione per il calcolo del totale camere
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_camere", "FUNCTION") + _
      " CREATE FUNCTION dbo.fn_calcola_totale_camere (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      " BEGIN " + vbCrLF + _
      "     /* " + vbCrLF + _
      "     Camere singole:			168, 342 " + vbCrLF + _
      "     Camere doppie:			169, 343 " + vbCrLF + _
      "     Camere a pi letti:		170, 344  " + vbCrLF + _
      "     Suite				172 " + vbCrLF + _
      "     Juniorsuite:			171 " + vbCrLF + _
      "     */  " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMERE int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMERE = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "         AND rel_grp_id_dotaz IN (168, 342, 169, 343, 170, 344, 172, 171) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Monolocali:		187 " + vbCrLF + _
      "     Bilocali:		188 " + vbCrLF + _
      "     Trilocali:		189 " + vbCrLF + _
      "     Pi locali:		337 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_UA int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (187, 188) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + SUM(rel_str_dotaz_valore * 2)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (189) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + SUM(rel_str_dotaz_valore - 1)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 3 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (337) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Totale Piazzole:	301 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMPEGGI int  " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = SUM(rel_str_dotaz_valore) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (301) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + SUM((num_vani - 1) * qta_ua)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)>1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + SUM(num_vani * qta_ua)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)=1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      vbCrLF + _
      "     RETURN ISNULL(@TOT_CAMERE_CAMERE, 0) + ISNULL(@TOT_CAMERE_UA, 0) + ISNULL(@TOT_CAMERE_CAMPEGGI, 0) " + vbCrLF + _
      " END "
CALL DB.Execute(sql, 1205)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1206
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__109(conn)
CALL DB.Execute(sql, 1206)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1207
'...........................................................................................
'non eseguito per conflitto con tb_servizi
'sql = Aggiornamento__FRAMEWORK_CORE__110(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1207)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1208
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__111(conn)
CALL DB.Execute(sql, 1208)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1209
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__112(conn)
CALL DB.Execute(sql, 1209)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1210
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__113(conn)
CALL DB.Execute(sql, 1210)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1211
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__114(DB, rs, 1211)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1212
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 1212)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1213
'riaggiorna la vista per dell'aggiornamento 115
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__115(conn)
CALL DB.Execute(sql, 1213)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1213)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1214
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__116(conn)
'	  Aggiornamento__FRAMEWORK_CORE__116_bis(conn):		non eseguito per conflitto con tb_servizi
CALL DB.Execute(sql, 1214)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1215
'...........................................................................................
'non eseguito per conflitto con tb_servizi
'sql = Aggiornamento__FRAMEWORK_CORE__117(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1215)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1216
'...........................................................................................
'non eseguito per conflitto con tb_servizi
'sql = Aggiornamento__FRAMEWORK_CORE__118(conn)
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1216)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1217
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__119(conn)
CALL DB.Execute(sql, 1217)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1218
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__120(conn)
CALL DB.Execute(sql, 1218)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1219
'...........................................................................................
CALL AggiornamentoSpeciale__FRAMEWORK_CORE__121(DB, rs, 1219)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1220
'...........................................................................................
'	corregge funzione per il calcolo del totale camere
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_camere", "FUNCTION") + _
      " CREATE FUNCTION dbo.fn_calcola_totale_camere (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      " BEGIN " + vbCrLF + _
      "     /* " + vbCrLF + _
      "     Camere singole:			168, 342 " + vbCrLF + _
      "     Camere doppie:			169, 343 " + vbCrLF + _
      "     Camere a pi letti:		170, 344  " + vbCrLF + _
      "     Suite				172 " + vbCrLF + _
      "     Juniorsuite:			171 " + vbCrLF + _
      "     */  " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMERE int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMERE = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "         AND rel_grp_id_dotaz IN (168, 342, 169, 343, 170, 344, 172, 171) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Monolocali:		187 " + vbCrLF + _
      "     Bilocali:		188 " + vbCrLF + _
      "     Trilocali:		189 " + vbCrLF + _
      "     Pi locali:		337 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_UA int " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)),0) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (187, 188) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + ISNULL(SUM(ISNULL(rel_str_dotaz_valore * 2, 0)), 0)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (189) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_UA = (@TOT_CAMERE_UA + ISNULL(SUM(ISNULL(rel_str_dotaz_valore - 1,0)), 0)) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 3 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (337) " + vbCrLF + _
      vbCrLF + _
      "     /* " + vbCrLF + _
      "     Totale Piazzole:	301 " + vbCrLF + _
      "     */ " + vbCrLF + _
      "     DECLARE @TOT_CAMERE_CAMPEGGI int  " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)),0) " + vbCrLF + _
      "         FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "         WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "               AND  rel_grp_id_dotaz IN (301) " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + ISNULL(SUM(ISNULL((num_vani - 1), 0) * ISNULL(qta_ua, 0)),0)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)>1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      "     SELECT @TOT_CAMERE_CAMPEGGI = (@TOT_CAMERE_CAMPEGGI + ISNULL(SUM(ISNULL(num_vani * qta_ua, 0)), 0)) " + vbCrLF + _
      "         FROM tb_ua WHERE id_struttura_ua = @STR_ID AND IsNull(num_vani,0)=1 AND ISNULL(qta_ua, 0)>0 " + vbCrLF + _
      vbCrLF + _
      "     RETURN ISNULL(@TOT_CAMERE_CAMERE, 0) + ISNULL(@TOT_CAMERE_UA, 0) + ISNULL(@TOT_CAMERE_CAMPEGGI, 0) " + vbCrLF + _
      " END "
CALL DB.Execute(sql, 1220)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1221
'...........................................................................................
'	aggiunge funzione per il calcolo del totale posti letto.
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_posti_letto", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_calcola_totale_posti_letto (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         /* " + vbCrLF + _
      "         Totale posti letto camere:		174 " + vbCrLF + _
      "         Totale posti letto suite:		176 " + vbCrLF + _
      "         Totale posti lestto u.a.:		195 " + vbCrLF + _
      "         CRM campeggi:				    308 " + vbCrLF + _
      "         */ " + vbCrLF + _
      vbCrLF + _
      "         DECLARE @TOT_PL int " + vbCrLF + _
      "         SELECT @TOT_PL = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0) " + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND  rel_grp_id_dotaz IN (174, 176, 195, 308) " + vbCrLF + _
      "         RETURN @TOT_PL " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 1221)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1222
'...........................................................................................
'	aggiorna funzione per il calcolo del totale bagni
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_bagni", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_calcola_totale_bagni (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      "  " + vbCrLF + _
      " AS  " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         /* " + vbCrLF + _
      "         Totale generale bagni:				207 " + vbCrLF + _
      "         Camerini bagno chiusi:				317 " + vbCrLF + _
      "         Servizi igienici per singoli equipaggi:		326 " + vbCrLF + _
      "         Servizi igienici per disabili:			318 " + vbCrLF + _
      "         Unit abitative con servizi igienici:		303 " + vbCrLF + _
      "         */ " + vbCrLF + _
      vbCrLF + _
      "         DECLARE @TOT_BAGNI int " + vbCrLF + _
      "         SELECT @TOT_BAGNI = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0)" + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND  rel_grp_id_dotaz IN (207, 317, 326, 318) " + vbCrLF + _
      vbCrLF + _
      "         SELECT @TOT_BAGNI = @TOT_BAGNI + ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0) " + vbCrLF + _
      "             FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "             WHERE rel_str_dotaz.rel_str_dotaz_pos_val IN (1, 3)  AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                   AND rel_grp_id_dotaz IN (303)  " + vbCrLF + _
      "         RETURN @TOT_BAGNI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 1222)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1223
'...........................................................................................
'	aggiunge funzione per il calcolo degli esercizi
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_esercizi", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_calcola_totale_esercizi (@STR_ID int, @MODELLO int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      " AS " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "         DECLARE @TOT_ESERCIZI int " + vbCrLF + _
      vbCrLF + _
      "         if (@MODELLO = 21 OR @MODELLO=31 OR @MODELLO=36) " + vbCrLF + _
      "         BEGIN " + vbCrLF + _
      "             SELECT @TOT_ESERCIZI = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0) " + vbCrLF + _
      "                 FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id " + vbCrLF + _
      "                 WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                       AND rel_grp_id_dotaz IN (194) " + vbCrLF + _
      "         END " + vbCrLF + _
      "         ELSE BEGIN " + vbCrLF + _
      "             SET @TOT_ESERCIZI = 1 " + vbCrLF + _
      "         END " + vbCrLF + _
      vbCrLF + _
      "         RETURN @TOT_ESERCIZI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 1223)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1224
'...........................................................................................
'	aggiorna funzione per il calcolo del totale bagni
'...........................................................................................
sql = DropObject(conn, "fn_calcola_totale_bagni", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_calcola_totale_bagni (@STR_ID int) " + vbCrLF + _
      "     RETURNS int " + vbCrLF + _
      "  " + vbCrLF + _
      " AS  " + vbCrLF + _
      "     BEGIN " + vbCrLF + _
      "     	/* " + vbCrLF + _
      "     	Totale generale bagni:				207 " + vbCrLF + _
      "     	Camerini bagno chiusi:				317 " + vbCrLF + _
      "     	Servizi igienici per disabili:			318 " + vbCrLF + _
      "     	Unit abitative con servizi igienici:		303 pos 1 e 3 " + vbCrLF + _
      "     	Servizi igienici per disabili:			305 pos 1 e 3 " + vbCrLF + _
      "     	Servizi igienici per singoli equipaggi		326 " + vbCrLF + _
      "     	Camerini wc inglesi				310 " + vbCrLF + _
      "     	Camerini wc turca				311 " + vbCrLF + _
      "     	*/  " + vbCrLF + _
	  vbcrlf + _
      "     	DECLARE @TOT_BAGNI int  " + vbCrLF + _
      "     	SELECT @TOT_BAGNI = ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0) " + vbCrLF + _
      "     		FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id  " + vbCrLF + _
      "     	WHERE rel_str_dotaz.rel_str_dotaz_pos_val = 1 AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID  " + vbCrLF + _
      "     		AND  rel_grp_id_dotaz IN (207, 317, 326, 318, 310, 311) " + vbCrLF + _
	  vbcrlf + _
      "     	SELECT @TOT_BAGNI = @TOT_BAGNI + ISNULL(SUM(ISNULL(rel_str_dotaz_valore,0)), 0)  " + vbCrLF + _
      "     		FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id  " + vbCrLF + _
      "     	WHERE rel_str_dotaz.rel_str_dotaz_pos_val IN (1, 3)  AND rel_str_dotaz.rel_id_str_dotaz = @STR_ID  " + vbCrLF + _
      "     		AND rel_grp_id_dotaz IN (303, 305) " + vbCrLF + _
	  vbcrlf + _
      "     	RETURN @TOT_BAGNI " + vbCrLF + _
      "     END "
CALL DB.Execute(sql, 1224)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1225
'...........................................................................................
'	26/03/2009 - Nicola
'...........................................................................................
'	modifica stored procedure per il calcolo del codice regionale
'...........................................................................................
sql = DropObject(conn, "fn_NEW_REGCODE", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_NEW_REGCODE(" + vbCrLF + _
	  "		@COMUNE nvarchar(6), " + vbCrLF + _
      "		@MODELLO int " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " RETURNS nvarchar(12) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " BEGIN " + vbCrLF + _
	  "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
	  "     DECLARE @FIRST_LETTER nvarchar(3) " + vbCrLF + _
	  "     DECLARE @VAR_CODEPART int " + vbCrLF + _
	  "     DECLARE @CODE_LENGHT int " + vbCrLF + _
      "     DECLARE @VAR_CODEPART_LENGTH int " + vbCrLF + _
      vbCrLF + _
      "     --recupera radicie del codice regionale " + vbCrLF + _
	  "     SELECT @FIRST_LETTER = Mod_FirstLT_Regcode FROM tb_Modelli WHERE Mod_ID= @MODELLO " + vbCrLF + _
      vbCrLF + _
      "     --compone prima parte del codice regionale " + vbCrLF + _
      "     SET @REGCODE = @FIRST_LETTER + @COMUNE " + vbCrLF + _
      vbCrLF + _
      "     --calcola lunghezza parte incrementale del codice " + vbCrLF + _
      "     IF (LEN(@FIRST_LETTER)) > 1 " + vbCrLF + _
      "         SET @CODE_LENGHT = 12 " + vbCrLF + _
      "     ELSE " + vbCrLF + _
	  "         SET @CODE_LENGHT = 11 " + vbCrLF + _
	  "     SET @VAR_CODEPART_LENGTH = @CODE_LENGHT - LEN(@REGCODE) " + vbCrLF + _
      vbCrLF + _
      "     --recupera parte incrementale del codice (ultimo inserito) " + vbCrLF + _
      "     SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(LTRIM(RTRIM(str_log_codAlb)), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
      "         FROM tb_str_logs " + vbCrLF + _
      "         WHERE LTRIM(RTRIM(str_log_codAlb)) LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "               LEN(LTRIM(RTRIM(str_log_codAlb))) = @CODE_LENGHT " + vbCrLF + _
      vbCrLF + _
      "     --incrementa parte variabile " + vbCrLF + _
      "     SET @VAR_CODEPART = @VAR_CODEPART + 1 " + vbCrLF + _
      vbCrLF + _
      "     --compone codice regionale definitivo " + vbCrLF + _
      "     SET @REGCODE = @REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART AS NVARCHAR(12)) " + vbCrLF + _
      vbCrLF + _
	  "     RETURN @REGCODE " + vbCrLF + _
	  " END " + vbCrLF
CALL DB.Execute(sql, 1225)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1226
'...........................................................................................
'	26/03/2009 - Nicola
'...........................................................................................
'	aggiunge campo per collegamento u.a.c. e u.a.n.c. con le rispettive u.a. gestite da agenzia
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD cod_ua_gestita nvarchar(12) NULL; "
CALL DB.Execute(sql, 1226)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1227
'...........................................................................................
'	31/03/2009 - Nicola
'...........................................................................................
'	aggiunge campo per gestioen scambio record su log per trasferimento dati regione ua gestite
'...........................................................................................
sql = " ALTER TABLE tb_str_logs ADD " + _
	  "		str_log_codalb_regione nvarchar(24) NULL, " + _
	  "		Str_log_record_regione nvarchar(60) NULL; " + _
	  " UPDATE tb_str_logs SET " + _
	  "		str_log_codalb_regione = str_log_codalb, " + _
	  "		Str_log_record_regione = str_log_record ; "
CALL DB.Execute(sql, 1227)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1227)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1228
'...........................................................................................
'	aggiorna trigger alla struttura di log per il calcolo del progressivo
'...........................................................................................
sql = DropObject(conn, "tb_str_logs_INSERT", "TRIGGER") + _
	  " CREATE TRIGGER dbo.tb_str_logs_INSERT ON tb_str_logs AFTER INSERT AS " + vbCrLf + _
      "     DECLARE @REGCODE nvarchar(13) " + vbCrLF + _
	  "     DECLARE @REGCODE_REGIONE nvarchar(13) " + vbCrLF + _
	  "		DECLARE @STR_ID nvarchar(60) " + vbCrlf + _
  	  "     DECLARE @STR_ID_REGIONE nvarchar(60) " + vbCrLF + _
      "     DECLARE @LAST INT " + vbCrLF + _
      "     DECLARE @PROGRESSIVO INT " + vbCrLF + _
      vbCrLf + _
      "     SELECT @REGCODE = RTRIM(LTRIM(str_log_CodAlb)), @LAST=str_log_id, @STR_ID = str_log_record " + vbCrLF + _
      "         FROM INSERTED WHERE ((Str_log_des LIKE '%registrazione validata%') OR (Str_log_des LIKE '%cancellazione completa struttura%')) " + vbCrLf + _
      vbCrLF + _
      "     if (@REGCODE <> '') " + vbCRLF + _
      "         BEGIN " + vbCrLF + _
	  "				SELECT @REGCODE_REGIONE = regcode, @STR_ID_REGIONE = str_id FROM view_strutture WHERE RTRIM(LTRIM(IsNull(cod_ua_gestita,''))) LIKE @REGCODE AND IsNull(gestito_agenzia,0)=1 " + vbCrLf + _
	  "				IF (IsNull(@REGCODE_REGIONE, '')='') " + vbCrLf + _
	  "					BEGIN " + vbCrLf + _
	  "						SET @REGCODE_REGIONE = @REGCODE " + vbCrLF + _
	  "						SET @STR_ID_REGIONE = @STR_ID " + vbCrLF + _
	  "					END " + vbCrLf + _
	  vbCrLF + _
      "             IF (EXISTS(SELECT * FROM tb_str_logs WHERE RTRIM(LTRIM(str_log_CodAlb_regione)) LIKE @REGCODE_REGIONE AND str_log_id <> @LAST AND IsNull(str_log_progressivo,0)<>0 )) " + vbCrLf + _
      "                 BEGIN " + vbCrLF + _
      "                     SELECT @PROGRESSIVO = MAX(str_log_progressivo) FROM tb_str_logs " + vbCrLF + _
      "                         WHERE RTRIM(LTRIM(str_log_CodAlb_regione)) LIKE @REGCODE_REGIONE AND str_log_id <> @LAST " + vbCrLf + _
      "                     SET @PROGRESSIVO = @PROGRESSIVO + 1 " + vbCrLF + _
      "                 END " + vbCrLf + _
      "             ELSE " + vbCrLF + _
      "                 BEGIN " + vbCrLF + _
      "                     SET @PROGRESSIVO = 1 " + vbCrLF + _
      "                 END " + vbCrLF + _
      "             UPDATE tb_str_logs SET str_log_codalb_regione = @REGCODE_REGIONE, str_log_record_regione = @STR_ID_REGIONE, str_log_progressivo=@PROGRESSIVO WHERE str_log_id=@LAST " + vbCrLF + _
      "         END " + vbCrLF + _
      " ; "
CALL DB.Execute(sql, 1228)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1229
'...........................................................................................
'	20/04/2009 - Nicola
'...........................................................................................
'	aggiunge colonna ai servizi ed alle dotazioni per collegamento ad ENIT
'...........................................................................................
sql = " ALTER TABLE tb_servizi ADD " + _
	  "		serv_sigla_enit nvarchar(2) NULL ; " + _
	  " ALTER TABLE tb_dotazioni ADD " + _
	  "		dotaz_sigla_enit nvarchar(2) NULL ;"
CALL DB.Execute(sql, 1229)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1230
'...........................................................................................
'	20/04/2009 - Nicola
'...........................................................................................
'	aggiunge funzione per la lettura dei servizi
'...........................................................................................
sql = " CREATE FUNCTION dbo.fn_ENIT_lista_servizi (@STR_ID int) " + vbCrLF + _
      "       RETURNS nvarchar(255) " + vbCrLF + _
      "    AS " + vbCrLF + _
      "    BEGIN" + vbCrLF + _
      "            DECLARE RS CURSOR " + vbCrLF + _
      "                FOR SELECT DISTINCT serv_sigla_enit " + vbCrLF + _
      "                    FROM tb_servizi INNER JOIN rel_Grp_serv ON tb_servizi.serv_id = rel_Grp_serv.rel_Grp_id_serv " + vbCrLF + _
      "                                    INNER JOIN rel_str_serv ON rel_Grp_serv.rel_Grp_serv_id = rel_str_serv.rel_str_id_relserv " + vbCrLF + _
      "                    WHERE rel_id_str_serv = @STR_ID " + vbCrLF + _
      "                          AND IsNull(serv_sigla_enit,'')<>'' " + vbCrLF + _
      "                    UNION " + vbCrLF + _
      "                    SELECT DISTINCT dotaz_sigla_enit " + vbCrLF + _
      "                    FROM dbo.tb_dotazioni INNER JOIN dbo.rel_Grp_dotaz ON dbo.tb_dotazioni.dotaz_id = dbo.rel_Grp_dotaz.rel_Grp_id_dotaz " + vbCrLF + _
      "                                          INNER JOIN dbo.rel_str_dotaz ON dbo.rel_Grp_dotaz.rel_Grp_dotaz_id = dbo.rel_str_dotaz.rel_str_id_dotaz " + vbCrLF + _
      "                    WHERE rel_id_str_dotaz = @STR_ID " + vbCrLF + _
      "                          AND IsNull(dotaz_sigla_enit,'')<>'' " + vbCrLF + _
      vbcrlf + _
      "            DECLARE @sigla nvarchar(2) " + vbCrLF + _
      "            DECLARE @LISTA_SERVIZI nvarchar(255) " + vbCrLF + _
      "            SET @LISTA_SERVIZI = ''" + vbCrLF + _
      "            OPEN RS" + vbCrLF + _
      "            FETCH NEXT FROM RS INTO @sigla " + vbCrLF + _
      "            WHILE (@@fetch_status <> -1) " + vbCrLF + _
      "                BEGIN" + vbCrLF + _
      "                    IF (@@fetch_status <> -2) " + vbCrLF + _
      "                        BEGIN " + vbCrLF + _
      "                            SET @LISTA_SERVIZI = @LISTA_SERVIZI + @sigla + ' ' " + vbCrLF + _
      "                        END " + vbCrLF + _
      "                    FETCH NEXT FROM RS INTO @sigla " + vbCrLF + _
      "                END " + vbCrLF + _
      "            CLOSE RS " + vbCrLF + _
      "            DEALLOCATE RS " + vbCrLF + _
      "            --colazione inclusa " + vbCrLF + _
      "            IF( (SELECT COUNT(*) FROM rel_str_dotaz WHERE rel_id_str_dotaz = @STR_ID AND rel_str_id_dotaz = 136 AND rel_str_dotaz_valore = 1 AND rel_str_dotaz_pos_val=1) >0 ) " + vbCrLF + _
      "                SET @LISTA_SERVIZI = @LISTA_SERVIZI + 'CO' " + vbCrLF + _
      "            RETURN @LISTA_SERVIZI " + vbCrLF + _
      "            END "
CALL DB.Execute(sql, 1230)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1231
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__122(conn)
CALL DB.Execute(sql, 1231)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1232
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__123(conn)
CALL DB.Execute(sql, 1232)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1233
'NICOLA 12/05/2009
'aggiorna sincronizzazione agenzie immobiliari
'...........................................................................................
sql = " SELECT * FROM AA_versione "
CALL DB.Execute(sql, 1233)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1233(DB.objConn, rs, rst, 34)
end if
CALL DB.ReSyncTransaction()
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub AggiornamentoSpeciale_1233(conn, rs, rst, MODELLO)
	dim sql, readConn, readRs
	'crea nuova connessione per evitare inferferenza con transazioni
	set readConn = Server.CreateObject("ADODB.Connection")
	set readRs = Server.CreateObject("ADODB.RecordSet")
	readConn.open conn.ConnectionString, "", ""
	readconn.CommandTimeout = 180
	
	sql = " SELECT modello, RegCode, mod_tipo_record FROM VIEW_Strutture " + _
		  " WHERE modello = " & modello
	readRs.open sql, readConn, adOpenStatic, adLockReadOnly, adCmdText
	while not readRs.eof %>
		<!-- <%= readRs("RegCode") %> - <%= readRs.absoluteposition %> su <%= readRs.recordcount %>-->
		<%CALL SincronizzaStruttura_NextCom_NextInfo(readConn, Conn, rs, readRs("modello"), readRs("RegCode"), readRs("mod_tipo_record"), true)
		readRs.movenext
	wend
	readRs.close
	
	readConn.close
	set readRs = nothing
	set readConn = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1234
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__124(conn)
CALL DB.Execute(sql, 1234)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1235
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__125(conn)
CALL DB.Execute(sql, 1235)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1236
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__126(conn)
CALL DB.Execute(sql, 1236)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1237
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__127(conn)
CALL DB.Execute(sql, 1237)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1238
'...........................................................................................
'	10/07/2009 - Nicola
'...........................................................................................
'	aggiunge colonne di gestione dichiarazione accesibilita' su modelli
'...........................................................................................
sql = " ALTER TABLE tb_modelli ADD " + _
	  "		mod_dichiarazione_accessibilita_attiva BIT NULL, " + _
	  "		mod_dichiarazione_accessibilita_obbligatoria BIT NULL, " + _
	  "		mod_dichiarazione_accessibilita_modulo nvarchar(250) NULL ; " + _
	  " UPDATE tb_modelli SET mod_dichiarazione_accessibilita_attiva = 1, " + _
	  "		mod_dichiarazione_accessibilita_obbligatoria = (CASE WHEN mod_categoria_max>0 THEN 1 ELSE 0 END), " + _
	  "		mod_dichiarazione_accessibilita_modulo = (CASE WHEN mod_id=18 THEN 'modulo_accessibilita_alberghiero.asp' ELSE 'modulo_accessibilita_extralberghiero.asp' END) " +_
	  "		WHERE mod_tipo_record IN ('S', 'O') ; "
CALL DB.Execute(sql, 1238)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1238)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1239
'...........................................................................................
'	10/07/2009 - Nicola
'...........................................................................................
'	aggiunge tabella per gestione dichiarazione accessibilita
'...........................................................................................
sql = " CREATE TABLE dbo.tb_str_dichiarazione_accessiblita ( " + _
	  "		da_id " + SQL_PrimaryKey(conn, "tb_str_dichiarazione_accessiblita") + "," + _
	  "		da_regcode nvarchar(12) NOT NULL, " + _
	  "		da_datamodifica smalldatetime NOT NULL, " + _
	  "		da_rapp_nominativo nvarchar(255), " + _
	  "		da_rapp_nato_a nvarchar(255) NULL, " + _
	  " 	da_rapp_nato_il nvarchar(255) NULL, " + _
	  "		da_rapp_comune nvarchar(255) NULL, " + _
	  "		da_rapp_indirizzo nvarchar(255) NULL, " + _
	  "		da_rapp_civico nvarchar(255) NULL, " + _
	  "		da_rapp_tipo nvarchar(255) NULL, " + _
	  "		da_accessibilita_totale BIT NULL, " + _
	  "		da_accessibilita_parziale BIT NULL, " + _
	  "		da_accessibilita_assente BIT NULL " + _
	  " ) " + _
	  SQL_AddForeignKey(conn, "tb_str_dichiarazione_accessiblita", "da_regcode", "tb_loginstru", "codAlb", true, "")
CALL DB.Execute(sql, 1239)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1240
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__128(conn)
CALL DB.Execute(sql, 1240)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1241
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__129(conn)
CALL DB.Execute(sql, 1241)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1242
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__130(conn)
CALL DB.Execute(sql, 1242)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1242)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1243
'Aggiunge dati per attivazione ambito Cavallino
'Nicola  01/02/2010
'...........................................................................................
sql = " INSERT INTO tb_apt (apt_codice, apt_nome) VALUES ('15', 'N.15 Cavallino') ; " + _
	  " UPDATE tb_apt_uffici SET uf_nome = 'N.15 Cavallino', uf_apt_codice = '15' WHERE uf_id = 09 ; " + _
	  " UPDATE tb_comuni SET cod_APT = '15' WHERE codice_ISTAT='027044'; " + _
	  " UPDATE tb_strutture SET AptCode='15' WHERE comune = '027044'; "
CALL DB.Execute(sql, 1243)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1243(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub AggiornamentoSpeciale_1243(conn, rs, rst)
	dim sql, readConn, readRs
	sql = " SELECT RegCode, mod_tipo_record, str_id FROM VIEW_Strutture " + _
		  " WHERE mod_tipo_record <> 'P' " + _
				" AND ( comune = '027044' OR RegCode IN (SELECT cod_proprietario FROM view_strutture where comune = '027044') " + _
										" OR RegCode IN (SELECT cod_tipologia FROM view_strutture where comune = '027044') ) " + _
		  " ORDER BY ( CASE WHEN mod_tipo_record LIKE 'S'  THEN 0 " + _
					 " WHEN mod_tipo_record LIKE 'U' THEN 1 " + _
					 " WHEN mod_tipo_record LIKE 'A' THEN 2 " + _
					 " WHEN mod_tipo_record LIKE 'O' THEN 3 " + _
					 " WHEN mod_tipo_record LIKE 'T' THEN 4 END) "
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		CALL SET_APT(conn, rst, rs("str_id"))
		rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1244
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__131(conn)
CALL DB.Execute(sql, 1244)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1245
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__132(conn)
CALL DB.Execute(sql, 1245)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1246
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__133(conn)
CALL DB.Execute(sql, 1246)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1246)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1247
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__134(conn)
CALL DB.Execute(sql, 1247)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1248
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__135(conn)
CALL DB.Execute(sql, 1248)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1249
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__136(conn)
CALL DB.Execute(sql, 1249)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1250
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__137(conn)
CALL DB.Execute(sql, 1250)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1250)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1251
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__138(conn)
CALL DB.Execute(sql, 1251)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1252
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__139(conn)
CALL DB.Execute(sql, 1252)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1253
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__140(conn)
CALL DB.Execute(sql, 1253)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1254
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__141(conn)
CALL DB.Execute(sql, 1254)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1254)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1255
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__142(conn)
CALL DB.Execute(sql, 1255)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1255)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1256
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__143(conn)
CALL DB.Execute(sql, 1256)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1256)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1257
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__144(conn)
CALL DB.Execute(sql, 1257)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1257)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1258
'
'Giacomo  09/12/2009
'...........................................................................................
sql =   "	ALTER TABLE tb_zone_urb ADD " + _
		"	zonaurb_cod_rvtweb " + SQL_CharField(Conn, 20) + " NULL;" + _
		"	ALTER TABLE rel_Grp_serv ADD " + _
		"	rel_cod_rvtweb " + SQL_CharField(Conn, 20) + " NULL;"
CALL DB.Execute(sql, 1258)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1258)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1259
'
'Giacomo  09/12/2009
'...........................................................................................
sql =   "	ALTER TABLE rel_Grp_dotaz ADD " + _
		"	rel_cod_rvtweb " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_sez_xml_rvtweb " + SQL_CharField(Conn, 20) + " NULL;"
CALL DB.Execute(sql, 1259)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1260
'
'Giacomo  18/12/2009
'...........................................................................................
sql =   "	ALTER TABLE rel_Grp_dotaz ADD " + _
		"	rel_cod_rvtweb_val_1 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_2 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_3 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_4 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_5 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_6 " + SQL_CharField(Conn, 20) + " NULL," + _
		"	rel_cod_rvtweb_val_7 " + SQL_CharField(Conn, 20) + " NULL "
CALL DB.Execute(sql, 1260)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1260)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1261
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__145(conn)
CALL DB.Execute(sql, 1261)
'*******************************************************************************************


'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 1262
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "ru")
CALL DB.Execute(sql, 1262)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "ru", "russo", "Русский")
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1263
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "cn")
CALL DB.Execute(sql, 1263)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "cn", "Cinese", "中文")
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1263)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1264
'...........................................................................................
'inserisce tutti i campi per l'aggiunta di una nuova lingua
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__146(conn, "pt")
CALL DB.Execute(sql, 1264)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, "pt", "Portoghese", "Português")
end if
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1264)
'*******************************************************************************************

'*******************************************************************************************
'chiusura transazione per liberare risorse bloccate
'...........................................................................................
CALL DB.ReSyncTransaction()
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1265
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__148(conn)
CALL DB.Execute(sql, 1265)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1266
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__149(conn)
CALL DB.Execute(sql, 1266)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1266)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1267
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__150(conn)
CALL DB.Execute(sql, 1267)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1267)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1268
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__151(conn)
CALL DB.Execute(sql, 1268)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1269
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__152(conn)
CALL DB.Execute(sql, 1269)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1270
'...........................................................................................
sql = Aggiornamento__INFO__21(conn)
CALL DB.Execute(sql, 1270)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1271
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__153(conn)
CALL DB.Execute(sql, 1271)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1271)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1272
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__154(conn)
CALL DB.Execute(sql, 1272)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1273
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__155(conn)
CALL DB.Execute(sql, 1273)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1274
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__156(conn)
CALL DB.Execute(sql, 1274)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1274)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1275
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__157(conn)
CALL DB.Execute(sql, 1275)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1275)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1276
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__158(conn)
CALL DB.Execute(sql, 1276)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1277
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__159(conn)
CALL DB.Execute(sql, 1277)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1277)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1278
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__160(conn)
CALL DB.Execute(sql, 1278)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1278)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1279
'Giacomo  28/06/2010
' Cancellazione anagrafiche senza corrispondenza tra le strutture
'...........................................................................................
sql = DropObject(conn, "tb_modelli_INSERT", "TRIGGER")  + vbCrLF + _
	  " CREATE TRIGGER [tb_modelli_INSERT] ON [dbo].[tb_modelli] AFTER INSERT AS " + vbCrLF + _
	  " INSERT INTO tb_rubriche (nome_rubrica, SyncroTable, SyncroFilterTable, SyncroFilterKey, locked_rubrica, rubrica_esterna) " + vbCrLF + _
	  "		SELECT CASE mod_tipo_record " + vbCrLF + _
	  "			   	WHEN 'O' THEN 'Proprietari - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   	WHEN 'A' THEN 'Agenzie - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   	WHEN 'P' THEN 'Prof. Tur. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   	WHEN 'U' THEN 'U. A. - ' + LOWER(mod_strutture)" + vbCrLF + _
	  "				WHEN 'C' THEN 'Congressuale - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   	ELSE 'Strutture - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   END ,  " + vbCrLF + _
	  "			   'view_strutture', 'tb_modelli', mod_id, 1, 1 FROM INSERTED WHERE NOT(mod_tipo_record LIKE 'T') " + vbCrLF + _
	  "	INSERT INTO tb_rel_gruppiRubriche (id_dellaRubrica, id_gruppo_assegnato) " + vbCrLF + _
	  "		SELECT @@IDENTITY, id_gruppo FROM tb_gruppi; " + _
	  DropObject(conn, "tb_modelli_UPDATE", "TRIGGER")  + vbCrLF + _
	  " CREATE TRIGGER [tb_modelli_UPDATE] ON [dbo].[tb_modelli] AFTER UPDATE AS " + vbCrLF + _
	  " 	DECLARE @TIPO_RECORD nvarchar(1) " + vbCrLF + _
	  " 	DECLARE @MOD_ID int " + vbCrLF + _
	  " 	SELECT TOP 1 @TIPO_RECORD = mod_tipo_record, @MOD_ID = mod_id FROM INSERTED " + vbCrLF + _
	  " IF @TIPO_RECORD = 'T' " + vbCrLF + _
	  "		BEGIN " + vbCrLF + _
	  "			IF UPDATE(mod_tipo_record) " + vbCrLF + _
	  "				BEGIN " + vbCrLF + _
	  "					UPDATE tb_rubriche SET locked_rubrica=0, rubrica_esterna=0 " + vbCrLF + _
	  "					WHERE SyncroFilterTable LIKE 'tb_modelli' AND SyncroFilterKey=@MOD_ID " + vbCrLF + _
	  "				END " + vbCrLF + _
	  "		END " + vbCrLF + _
	  " ELSE " + vbCrLF + _
	  "		BEGIN " + vbCrLF + _
	  "			IF UPDATE(mod_strutture) " + vbCrLF + _
	  "				BEGIN " + vbCrLF + _
	  "					UPDATE tb_rubriche SET nome_rubrica = (SELECT TOP 1 CASE mod_tipo_record " + vbCrLF + _
	  "			   																WHEN 'O' THEN 'Proprietari - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   																WHEN 'A' THEN 'Agenzie - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "			   																WHEN 'P' THEN 'Prof. Tur. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				   															WHEN 'U' THEN 'U. A. - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "																			WHEN 'C' THEN 'Congressuale - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				   															ELSE 'Strutture - ' + LOWER(mod_strutture) " + vbCrLF + _
	  "				   															END " + vbCrLF + _
	  "																			FROM INSERTED) " + vbCrLF + _
	  "					WHERE SyncroFilterTable='tb_modelli' AND SyncroFilterKey  = @MOD_ID " + vbCrLF + _
	  "				END " + vbCrLF + _
	  "		END "
CALL DB.Execute(sql, 1279)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1280
'Aggiunta applicativi per gestione centri congressi
'Nicola  07/06/2010
'...........................................................................................
sql =   " INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_ALBERGHI & ", " & _
							  " 'Turismo Congressuale [Alberghi]', " & _
							  " '../admin/congressuale_alberghi', " & _
							  " 'CON_ALBERGHI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_CENTRICONGRESSI & ", " & _
							  " 'Turismo Congressuale [Centri congressi]', " & _
							  " '../admin/congressuale_centricongressi', " & _
							  " 'CON_CENTRICONGRESSI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_ALTRESEDI & ", " & _
							  " 'Turismo Congressuale [Altre sedi]', " & _
							  " '../admin/congressuale_altresedi', " & _
							  " 'CON_ALTRESEDI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_SEDISTORICHE & ", " & _
							  " 'Turismo Congressuale [Sedi storiche]', " & _
							  " '../admin/congressuale_sedistoriche', " & _
							  " 'CON_SEDISTORICHE_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_AGENZIE & ", " & _
							  " 'Turismo Congressuale [Agenzie con reparto congressuale]', " & _
							  " '../admin/congressuale_agenzie', " & _
							  " 'CON_AGENZIE_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_ORGANIZZATORI & ", " & _
							  " 'Turismo Congressuale [Organizzatori professionali]', " & _
							  " '../admin/congressuale_organizzatori', " & _
							  " 'CON_ORGANIZZATORI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_SERVIZITRADUZIONE & ", " & _
							  " 'Turismo Congressuale [Imprese di servizi traduzione ed interpretariato]', " & _
							  " '../admin/congressuale_servizitraduzione', " & _
							  " 'CON_SERVIZITRADUZIONE_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_SERVIZITECNICI & ", " & _
							  " 'Turismo Congressuale [Imprese di servizi tecnici]', " & _
							  " '../admin/congressuale_servizitecnici', " & _
							  " 'CON_SERVIZITECNICI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_IMPRESEASSISTENZA & ", " & _
							  " 'Turismo Congressuale [Imprese di assistenza congressuale]', " & _
							  " '../admin/congressuale_impreseassistenza', " & _
							  " 'CON_IMPRESEASSISTENZA_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbcrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_SERVIZICATERING & ", " & _
							  " 'Turismo Congressuale [Imprese servizi di catering]', " & _
							  " '../admin/congressuale_servizicatering', " & _
							  " 'CON_SERVIZICATERING_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_SERVIZITRASPORTO & ", " & _
							  " 'Turismo Congressuale [Imprese servizi di trasporto]', " & _
							  " '../admin/congressuale_servizitrasporto', " & _
							  " 'CON_SERVIZITRASPORTI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbcrLF + _
		" INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_CONGRESSUALE_ALLESTITORI & ", " & _
							  " 'Turismo Congressuale [Imprese servizi di allestimento]', " & _
							  " '../admin/congressuale_allestitori', " & _
							  " 'CON_ALLESTITORI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + _
		" ; " + _
		" INSERT INTO rel_admin_sito (admin_id, sito_id, rel_as_permesso) " + _
		" SELECT id_admin, id_sito, 1 " + _
		"	FROM tb_admin, tb_siti " + _
		"	WHERE id_admin IN (1,2) AND id_sito IN (" & TURISMO_CONGRESSUALE_ALBERGHI & ", " & _
														TURISMO_CONGRESSUALE_CENTRICONGRESSI & ", " & _
														TURISMO_CONGRESSUALE_ALTRESEDI & ", " & _
														TURISMO_CONGRESSUALE_SEDISTORICHE & ", " & _
														TURISMO_CONGRESSUALE_AGENZIE & ", " & _
														TURISMO_CONGRESSUALE_ORGANIZZATORI & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITRADUZIONE & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITECNICI & ", " & _
														TURISMO_CONGRESSUALE_IMPRESEASSISTENZA & ", " & _
														TURISMO_CONGRESSUALE_SERVIZICATERING & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITRASPORTO & ", " & _
														TURISMO_CONGRESSUALE_ALLESTITORI & " ) " + vbCrLf + _
		" INSERT INTO tb_turismo_admin_sito (tas_admin_id, tas_sito_id, tas_permesso, tas_ricezione_notifiche ) " + _
		" SELECT id_admin, id_sito, 4, 1 " + _
		"	FROM tb_admin, tb_siti " + _
		"	WHERE id_admin IN (1,2) AND id_sito IN (" & TURISMO_CONGRESSUALE_ALBERGHI & ", " & _
														TURISMO_CONGRESSUALE_CENTRICONGRESSI & ", " & _
														TURISMO_CONGRESSUALE_ALTRESEDI & ", " & _
														TURISMO_CONGRESSUALE_SEDISTORICHE & ", " & _
														TURISMO_CONGRESSUALE_AGENZIE & ", " & _
														TURISMO_CONGRESSUALE_ORGANIZZATORI & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITRADUZIONE & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITECNICI & ", " & _
														TURISMO_CONGRESSUALE_IMPRESEASSISTENZA & ", " & _
														TURISMO_CONGRESSUALE_SERVIZICATERING & ", " & _
														TURISMO_CONGRESSUALE_SERVIZITRASPORTO & ", " & _
														TURISMO_CONGRESSUALE_ALLESTITORI & " ) " + _
		" ; "  + _
		SQLSERVER_ReseedIdentity(conn, "tb_modelli", "mod_id") + _
		SQLSERVER_ReseedIdentity(conn, "tb_tipi_str", "tip_id") + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 			mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Alberghi con attività congressuale', 				'Alberghi con attività congressuale',	6,				   'CO',				'congressuale_alberghi',			'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_BY_STAR & "', 	0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_ALBERGHI & ", 			0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 			mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Centri congressi', 									'Centri congressi',						0,				   'CO',				'congressuale_centricongressi',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_CENTRICONGRESSI & ", 	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Altre sedi congressuali', 							'Altre sedi',							0,				   'CO',				'congressuale_altresedi',			'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_ALTRESEDI & ", 		0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Sedi congressuali storiche', 						'Sedi congressuali storiche',			0,				   'CO',				'congressuale_sedistoriche',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_SEDISTORICHE & ", 		0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Agenzie con reparto congressuale', 					'Agenzie congressi',					0,				   'CO',				'congressuale_agenzie',				'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_AGENZIE & ", 			0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Organizzatori professionali', 						'Organizzatori professionali',			0,				   'CO',				'congressuale_organizzatori',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_ORGANIZZATORI & ",		0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese di servizi traduzione ed interpretariato', 	'Servizi traduzione',					0,				   'CO',				'congressuale_servizitraduzione',	'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_SERVIZITRADUZIONE & ",	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese di servizi tecnici', 						'Servizi tecnici',						0,				   'CO',				'congressuale_servizitecnici',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_SERVIZITECNICI & ",	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese di assistenza congressuale', 				'Assistenza congressuale',				0,				   'CO',				'congressuale_servizitecnici',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_IMPRESEASSISTENZA & ",	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese servizi di catering', 						'Servizi di catering',					0,				   'CO',				'congressuale_servizicatering',		'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_SERVIZICATERING & ",	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese servizi di trasporto', 						'Servizi di trasporto',					0,				   'CO',				'congressuale_servizitrasporto',	'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_SERVIZITRASPORTO & ",	0					 )	" + vbCrLF + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 											Mod_strutture, 							Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 						mod_tipo_record, 		mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica ) " + _
						" VALUES ( 'Imprese servizi di allestimento', 					'Servizi di allestimento',				0,				   'CO',				'congressuale_allestitori',			'" & RECORD_TYPE_CC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_ALLESTITORI & 		",	0					 )	" + _
		" INSERT INTO tb_tipi_str (tip_mod_id, tip_den_it, tip_valid_from, tip_cod_regione, tip_cod_rvt) " + _
		"	SELECT mod_id, mod_strutture, CONVERT(DATETIME, '1990-01-01 00:00:00', 102), 0, 0 " + _
		"		FROM tb_modelli " + _
		"		WHERE mod_tipo_record LIKE '" & RECORD_TYPE_CC & "' " + _
		"; " + _
		" UPDATE tb_modelli SET Mod_FirstLT_Regcode = 'TC' WHERE mod_tipo_record LIKE '" & RECORD_TYPE_CC & "' "
CALL DB.Execute(sql, 1280)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1281
'Aggiunta applicativo area riservata
'Nicola  28/06/2010
'...........................................................................................
sql =   " INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " 100, " & _
							  " 'Area riservata operatori', " & _
							  " '', " & _
							  " 'USER', " & _
							  " 0, " & _
							  " '', " & _
							  " '' " & _
		" ) " + vbCrLF + _
		" INSERT INTO tb_rubriche (nome_rubrica, locked_rubrica, rubrica_esterna ) " + _
						" VALUES ('Utenti - Area riservata operatori', 1, 1) " + vbCrLF + _
		" UPDATE tb_siti SET sito_rubrica_area_riservata = @@IDENTITY WHERE id_sito = 100 " + vbcRlf + _
		" INSERT INTO tb_rel_gruppirubriche(id_dellaRubrica, id_gruppo_assegnato) " + _
		" SELECT top 1 sito_rubrica_area_riservata, id_Gruppo " + _
			   " FROM tb_siti, tb_gruppi WHERE id_sito = 100 "
CALL DB.Execute(sql, 1281)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1282
'Ricreazione stored procedure DELETE_Struttura
'Giacomo  28/06/2010
'...........................................................................................
sql = DropObject(conn, "DELETE_Struttura", "PROCEDURE") + vbCrLF + _
	  " CREATE Procedure [dbo].[DELETE_Struttura] " + vbCrLF + _
	  "  @RegCode nvarchar(12) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  "  DELETE FROM rel_zoneurb_str WHERE rel_str_id IN (SELECT str_ID FROM tb_Strutture WHERE regCode=@RegCode) " + vbCrLF + _
	  "  DELETE FROM rel_str_serv WHERE rel_id_str_serv IN (SELECT str_ID FROM tb_Strutture WHERE regCode=@RegCode) " + vbCrLF + _
	  "  DELETE FROM rel_str_dotaz WHERE rel_id_str_dotaz IN (SELECT str_ID FROM tb_Strutture WHERE regCode=@RegCode) " + vbCrLF + _
	  "  DELETE FROM tb_ua WHERE id_struttura_ua IN (SELECT str_ID FROM tb_Strutture WHERE regCode=@RegCode) " + vbCrLF + _
	  "  DELETE FROM tb_stru_gest WHERE regCode=@RegCode " + vbCrLF + _
	  "  DELETE FROM rel_assoc_stru WHERE CodAlb_rel = @RegCode " + vbCrLF + _
	  "  DELETE FROM tb_strutture WHERE RegCode = @RegCode " + vbCrLF + _
	  "  DELETE FROM tb_loginstru WHERE CodAlb = @RegCode " + vbCrLF + _
	  "  DELETE FROM itb_anagrafiche WHERE ana_id IN (SELECT IDElencoIndirizzi FROM tb_Indirizzario WHERE SyncroKey = @RegCode AND SyncroTable LIKE 'VIEW_valid_strutture') " + vbCrLF + _
	  "  DELETE FROM tb_Indirizzario WHERE SyncroKey = @RegCode AND SyncroTable LIKE 'VIEW_valid_strutture' "
CALL DB.Execute(sql, 1282)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1283
'Giacomo  29/06/2010
' Query di ripulitura
'...........................................................................................
sql = " UPDATE tb_loginstru SET codAlb = RTRIM(LTRIM(CodAlb)) " + vbCrLF + _
	  " UPDATE tb_loginStru SET CodAlb_Old = RTRIM(LTRIM(CodAlb_Old)) " + vbCrLF + _
	  " UPDATE tb_loginStru SET CodAlb_Regione = RTRIM(LTRIM(CodAlb_Regione)) " + vbCrLF + _
	  " UPDATE itb_anagrafiche SET ana_codice = RTRIM(LTRIM(ana_codice)) " + vbCrLF + _
	  " UPDATE tb_Indirizzario SET SyncroKey = RTRIM(LTRIM(SyncroKey)) " + vbCrLF + _
	  " UPDATE tb_str_dichiarazione_accessiblita SET da_regcode = RTRIM(LTRIM(da_regcode)) " + vbCrLF + _
	  " UPDATE tb_Str_logs SET Str_log_CodAlb = RTRIM(LTRIM(Str_log_CodAlb)) " + vbCrLF + _
	  " UPDATE tb_Str_logs SET str_log_codalb_regione = RTRIM(LTRIM(str_log_codalb_regione)) " + vbCrLF + _
	  " UPDATE tb_stru_gest SET RegCode = RTRIM(LTRIM(RegCode)) " + vbCrLF + _
	  " UPDATE tb_strutture SET RegCode = RTRIM(LTRIM(RegCode)) " + vbCrLF + _
	  " UPDATE tb_strutture SET Cod_Proprietario = RTRIM(LTRIM(Cod_Proprietario)) " + vbCrLF + _
	  " UPDATE tb_strutture SET Cod_Tipologia = RTRIM(LTRIM(Cod_Tipologia)) " + vbCrLF + _
	  " UPDATE tb_strutture SET cod_ua_gestita = RTRIM(LTRIM(cod_ua_gestita)) " + vbCrLF + _
	  " UPDATE rel_assoc_stru SET CodAlb_rel = RTRIM(LTRIM(CodAlb_rel)) "
CALL DB.Execute(sql, 1283)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1284
'Giacomo  28/06/2010
' Cancellazione anagrafiche senza corrispondenza tra le strutture
'...........................................................................................
sql = " DELETE FROM itb_anagrafiche WHERE ana_id IN (SELECT IDElencoIndirizzi FROM tb_indirizzario " + vbCrLF + _
	  " WHERE SyncroTable LIKE 'VIEW_valid_strutture' AND SyncroKey <> '' " + vbCrLF + _
	  " AND SyncroKey NOT IN (SELECT CODALB FROM tb_loginStru)); " + _
	  " DELETE FROM tb_indirizzario " + vbCrLF + _
	  " WHERE SyncroTable LIKE 'VIEW_valid_strutture' AND SyncroKey <> '' " + vbCrLF + _
	  " AND SyncroKey NOT IN (SELECT CODALB FROM tb_loginStru) "
CALL DB.Execute(sql, 1284)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1285
'Giacomo  28/06/2010
' Sincronizzazione del next-Passport con le strutture validate
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1285)
if DB.last_update_executed then
	CALL Aggiornamento_1285_SincronizzazioneStruttureNextPassport(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1285_SincronizzazioneStruttureNextPassport(DbConn, rs)
	dim objContatto, syncroID
	sql = "SELECT * FROM VIEW_valid_strutture WHERE mod_tipo_record NOT LIKE '" & RECORD_TYPE_TY & "' AND mod_tipo_record NOT LIKE '" & RECORD_TYPE_UA & "'"
	rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	while not rs.eof
		if cString(rs("login")) <> "" then
			syncroID = cIntero(GetValueList(conn, NULL, " SELECT IDElencoIndirizzi FROM tb_Indirizzario WHERE SyncroKey LIKE '" & rs("RegCode") & "' AND SyncroTable LIKE 'VIEW_valid_strutture'"))
			if syncroID > 0 then
				set objContatto = new IndirizzarioLock
				set objContatto.conn = conn
				CALL objContatto.LoadFromDB(syncroID)
				objContatto("login") = rs("LOGIN")
				objContatto("password") = rs("password")
				objContatto("Abilitato") = rs("struttura_attiva")
				CALL objContatto.UserFromContact(syncroID, SITO_AREA_RISERVATA)
				objContatto.conn = null
				set objContatto = nothing
			else
				response.write "ANAGRAFICA NON TROVATA! --> " & rs("Denominazione") & "<br>"
				response.write "RegCode:" & rs("RegCode") & "<br>"
				response.end
			end if
		end if
		rs.moveNext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1286
'Giacomo  29/06/2010
' Sincronizzazione del next-Passport con le strutture validate
'...........................................................................................
sql = "SELECT * FROM AA_versione"
CALL DB.Execute(sql, 1286)
if DB.last_update_executed then
	CALL Aggiornamento_1286_SincronizzazioneStruttureNextPassport(DB.objConn, rs)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub Aggiornamento_1286_SincronizzazioneStruttureNextPassport(DbConn, rs)
	dim objContatto, syncroID, conta
	conta = 1
	sql = "SELECT * FROM VIEW_valid_strutture WHERE mod_tipo_record NOT LIKE '" & RECORD_TYPE_TY & "' AND mod_tipo_record NOT LIKE '" & RECORD_TYPE_UA & "' AND [LOGIN]IS NULL AND PASSWORD IS NULL"
	rs.open sql, conn, adOpenDynamic, adLockOptimistic, adCmdText
	
	while not rs.eof
		syncroID = cIntero(GetValueList(conn, NULL, " SELECT IDElencoIndirizzi FROM tb_Indirizzario WHERE SyncroKey LIKE '" & rs("RegCode") & "' AND SyncroTable LIKE 'VIEW_valid_strutture'"))
		if syncroID > 0 then
			set objContatto = new IndirizzarioLock
			set objContatto.conn = conn
			CALL objContatto.LoadFromDB(syncroID)	
			if cString(rs("LOGIN")) <> "" then
				objContatto("login") = rs("LOGIN")
			else
				objContatto("login") = RANDOM_LOGIN_E_PASSWORD
			end if
			if cString(rs("password")) <> "" then
				objContatto("password") = rs("password")
			else
				objContatto("password") = RANDOM_LOGIN_E_PASSWORD
			end if
			objContatto("Abilitato") = rs("struttura_attiva")
			CALL objContatto.UserFromContact(syncroID, SITO_AREA_RISERVATA)
			
			if cString(rs("LOGIN")) = "" then
				DbConn.execute(" UPDATE tb_loginStru SET LOGIN = '" & objContatto("login") & "' WHERE CODALB LIKE '" & rs("RegCode") & "'")
			end if
			if cString(rs("password")) = "" then
				DbConn.execute(" UPDATE tb_loginStru SET PASSWORD = '" & objContatto("password") & "' WHERE CODALB LIKE '" & rs("RegCode") & "'")
			end if			
			
			objContatto.conn = null
			set objContatto = nothing
		else
			response.write conta & " ANAGRAFICA NON TROVATA! --> " & rs("Denominazione") & "<br>"
			response.write "RegCode:" & rs("RegCode") & "<br><br>"
			conta = conta + 1
			'response.end
		end if
		rs.moveNext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1287
'Giacomo  30/06/2010
' Aggiunta campo a tb_loginStru per estendere una struttura già esistente come centro congresso 
'...........................................................................................
sql = " ALTER TABLE tb_loginStru ADD CodAlb_Replicato nvarchar(12) NULL; "
CALL DB.Execute(sql, 1287)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1287)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1288
'Aggiunta applicativo e modelli e dati per agriturismi
'Nicola  01/07/2010
'...........................................................................................
sql =   " INSERT INTO tb_siti ( " + _
							  " id_sito," + _
							  " sito_nome," + _
							  " sito_dir," + _
							  " sito_p1," + _
							  " sito_amministrazione," + _
							  " sito_prmEsterni_admin," + _
							  " sito_prmEsterni_sito) " + _
		" VALUES (" +_
							  " " & TURISMO_AGRITURISMI & ", " & _
							  " 'Strutture ricettive [Agriturismi]', " & _
							  " '../admin/agriturismi', " & _
							  " 'AGRI_USER', " & _
							  " 1, " & _
							  " '../../Admin/Passport/PassportAdmin.asp', " & _
							  " '../../Admin/Passport/PassportSito.asp' " & _
		" ) " + vbCrLF + _
		" INSERT INTO rel_admin_sito (admin_id, sito_id, rel_as_permesso) " + _
		" SELECT id_admin, id_sito, 1 " + _
		"	FROM tb_admin, tb_siti " + _
		"	WHERE id_admin IN (1,2) AND id_sito = " & TURISMO_AGRITURISMI & vbCrLf + _
		" INSERT INTO tb_turismo_admin_sito (tas_admin_id, tas_sito_id, tas_permesso, tas_ricezione_notifiche ) " + _
		" SELECT id_admin, id_sito, 4, 1 " + _
		"	FROM tb_admin, tb_siti " + _
		"	WHERE id_admin IN (1,2) AND id_sito = " & TURISMO_AGRITURISMI & vbCrLf + _
		" ; "  + _
		SQLSERVER_ReseedIdentity(conn, "tb_modelli", "mod_id") + _
		SQLSERVER_ReseedIdentity(conn, "tb_tipi_str", "tip_id") + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 					Mod_strutture,	Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 			mod_tipo_record, 			mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 					mod_portale_pubblica, mod_rvtweb_codice ) " + _
						" VALUES ( 'Agriturismi', 				'Agriturismi',	0,				   'Y',					'agriturismi',			'" & RECORD_TYPE_STR & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_AGRITURISMI & ", 			0					, 'Y')	" + vbCrLF + _
		" INSERT INTO tb_tipi_str (tip_mod_id, tip_den_it, tip_valid_from, tip_cod_regione, tip_cod_rvt) " + _
		"	SELECT mod_id, mod_strutture, CONVERT(DATETIME, '1990-01-01 00:00:00', 102), 0, 0 " + _
		"		FROM tb_modelli " + _
		"		WHERE Mod_FirstLT_Regcode LIKE 'Y' " + _
		"; "
CALL DB.Execute(sql, 1288)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1289
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__161(conn)
CALL DB.Execute(sql, 1289)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1290
'Aggiunta modello per gestione delle sale congressi
'Nicola  07/07/2010
'...........................................................................................
sql =   SQLSERVER_ReseedIdentity(conn, "tb_modelli", "mod_id") + _
		SQLSERVER_ReseedIdentity(conn, "tb_tipi_str", "tip_id") + _
		" INSERT INTO tb_modelli ( Mod_Den_it, 					Mod_strutture,		Mod_Categoria_MAX, Mod_FirstLT_Regcode, mod_Directory, 			mod_tipo_record, 			mod_tipo_classificazione, 		mod_tabella_prezzi, mod_dichiarazione_online, mod_dichiarazione_tipo, 			mod_applicazione_id, 							mod_portale_pubblica, mod_rvtweb_codice ) " + _
						" VALUES ( 'Sale congressi', 			'Sale congressi',	0,				   'SC',				'congressuale_sale',	'" & RECORD_TYPE_SC & "', 	'" & CLASSIFIED_NONE & "', 		0,					0,						  '" & DICHIARAZIONE_NESSUNA & "',	" & TURISMO_CONGRESSUALE_CENTRICONGRESSI & ", 	0					, '')	" + vbCrLF + _
		" INSERT INTO tb_tipi_str (tip_mod_id, tip_den_it, tip_valid_from, tip_cod_regione, tip_cod_rvt) " + _
		"	SELECT mod_id, mod_strutture, CONVERT(DATETIME, '1990-01-01 00:00:00', 102), 0, 0 " + _
		"		FROM tb_modelli " + _
		"		WHERE mod_tipo_record LIKE '" & RECORD_TYPE_SC & "' " + _
		"; "
CALL DB.Execute(sql, 1290)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1291
'correzione tipologia di record per il turismo congressuale
'Nicola  07/07/2010
'...........................................................................................
sql = "	UPDATE tb_modelli SET mod_tipo_record = '" + RECORD_TYPE_CC + "' " + _
	  " WHERE mod_id IN (40, 41, 42, 43) ; " + _
	  "	UPDATE tb_modelli SET mod_tipo_record = '" + RECORD_TYPE_SC + "' " + _
	  " WHERE mod_id IN (53) ; " + _
	  "	UPDATE tb_modelli SET mod_tipo_record = '" + RECORD_TYPE_OC + "' " + _
	  " WHERE mod_id IN (44, 45, 46, 47, 48, 49, 50, 51) ; "
CALL DB.Execute(sql, 1291)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1292
'inserimento moduli di stampa per tutti i modelli del turismo congressuale
'Nicola  12/07/2010
'...........................................................................................
sql = " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (44,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_agenzie/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (44,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (40,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_alberghi/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (40,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (51,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_allestitori/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (51,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (42,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_altresedi/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (42,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (41,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_centricongressi/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (41,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (48,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_impreseassistenza/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (48,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (45,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_organizzatori/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (45,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (43,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_sedistoriche/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (43,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (49,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_servizicatering/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (49,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (47,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_servizitecnici/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (47,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (46,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_servizitraduzione/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (46,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (50,			  '" & PRINT_POS_SCHEDA & "', 	'Modello dichiarazione', 		0,					  1,					'', 			'../../riservata/congressuale_servizitrasporto/Modulo_Stampa.asp' ) ; " + vbCrLf + _
	  " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione, 			Stampa_nome,			 		Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " VALUES				  (50,			  '" & PRINT_POS_REPORT & "', 	'Elenco modelli dichiarazione', 1,					  0,					'',				'../../riservata/stampe/elenco_stampe.asp?Type=M' ) ; "
CALL DB.Execute(sql, 1292)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1293
'inserimento moduli di stampa vuoti per tutti i modelli del turismo congressuale
'Nicola  13/07/2010
'...........................................................................................
sql = " INSERT INTO tb_stampe (Stampa_modello, Stampa_posizione,          Stampa_nome,			 		 Stampa_param_modello, Stampa_param_RegCode, Stampa_note, 	Stampa_script_name) " + vbCrLf + _
	  " SELECT     			   Stampa_modello, '" + PRINT_POS_TUTTE + "', 'Modello dichiarazione vuoto', 0,					   0,					'',				REPLACE(Stampa_script_name, 'Modulo_Stampa', 'Modulo_Stampa_vuoto') " + vbCrLF + _
							 " FROM tb_stampe INNER JOIN tb_modelli ON tb_stampe.Stampa_modello = tb_modelli.Mod_ID " + _
							 " WHERE Stampa_posizione LIKE '" & PRINT_POS_SCHEDA & "' AND (mod_tipo_record LIKE '" & RECORD_TYPE_CC & "' OR mod_tipo_record LIKE '" & RECORD_TYPE_SC & "' OR mod_tipo_record LIKE '" & RECORD_TYPE_OC & "') "
CALL DB.Execute(sql, 1293)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 1294
'Creo le tabelle di foto e tipizzazione foto
'Giacomo  20/07/2010
'...........................................................................................
sql = 	"CREATE TABLE " & SQL_Dbo(Conn) & "tb_strutture_foto ("& _
			  "		foto_ID " & SQL_PrimaryKey(conn, "tb_strutture_foto") + ", "& _
			  "		RegCode " + SQL_CharField(Conn, 12) + " NOT NULL, "& _
			  "		foto_thumb " + SQL_CharField(Conn, 255) + " NULL, "& _
			  "		foto_zoom " + SQL_CharField(Conn, 255) + " NULL, "& _
			  "		foto_didascalia_it " + SQL_CharField(Conn, 0) + " NULL, " + _
			  "		foto_ordine INTEGER NULL, "& _
			  "		foto_tipo_id INTEGER NULL, "& _
			  "		foto_data_inserimento DATETIME NULL,"& _
			  "		foto_data_modifica DATETIME NULL"& _
			  ");"& _
			  SQL_AddForeignKey(conn, "tb_strutture_foto", "RegCode", "tb_loginStru", "CodAlb", true, "") + _
		" CREATE TABLE " & SQL_Dbo(conn) & "tb_strutture_foto_tipo ( " & _
			"	ft_id " + SQL_PrimaryKey(conn, "tb_strutture_foto_tipo") + ", " + _
			"	ft_nome " + SQL_CharField(Conn, 255) + " NULL, "+ _
			"	ft_codice " + SQL_CharField(Conn, 255) + " NULL " + _
			" ) ; " + _
			SQL_AddForeignKey(conn, "tb_strutture_foto", "foto_tipo_id", "tb_strutture_foto_tipo", "ft_id", true, "") + _
			" INSERT INTO tb_strutture_foto_tipo(ft_nome,ft_codice) VALUES ('immagini', 'img') ; " + _
			" INSERT INTO tb_strutture_foto_tipo(ft_nome,ft_codice) VALUES ('planimetrie', 'pln') ; " + _
			" UPDATE tb_strutture_foto SET foto_tipo_id = 1 ; "
CALL DB.Execute(sql, 1294)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1295
'...........................................................................................
'	03/08/2010 - Nicola
'...........................................................................................
'	modifica stored procedure per il calcolo del codice regionale
'...........................................................................................
sql = DropObject(conn, "fn_NEW_REGCODE", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_NEW_REGCODE(" + vbCrLF + _
	  "		@COMUNE nvarchar(6), " + vbCrLF + _
      "		@MODELLO int " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " RETURNS nvarchar(12) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " BEGIN " + vbCrLF + _
	  "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
	  "     DECLARE @FIRST_LETTER nvarchar(3) " + vbCrLF + _
	  "     DECLARE @VAR_CODEPART int " + vbCrLF + _
	  "     DECLARE @CODE_LENGHT int " + vbCrLF + _
      "     DECLARE @VAR_CODEPART_LENGTH int " + vbCrLF + _
      vbCrLF + _
      "     --recupera radicie del codice regionale " + vbCrLF + _
	  "     SELECT @FIRST_LETTER = Mod_FirstLT_Regcode FROM tb_Modelli WHERE Mod_ID= @MODELLO " + vbCrLF + _
      vbCrLF + _
      "     --compone prima parte del codice regionale " + vbCrLF + _
      "     SET @REGCODE = @FIRST_LETTER + @COMUNE " + vbCrLF + _
      vbCrLF + _
      "     --calcola lunghezza parte incrementale del codice " + vbCrLF + _
      "     IF (LEN(@FIRST_LETTER)) > 1 " + vbCrLF + _
      "         SET @CODE_LENGHT = 12 " + vbCrLF + _
      "     ELSE " + vbCrLF + _
	  "         SET @CODE_LENGHT = 11 " + vbCrLF + _
	  "     SET @VAR_CODEPART_LENGTH = @CODE_LENGHT - LEN(@REGCODE) " + vbCrLF + _
      vbCrLF + _
      "     --recupera parte incrementale del codice (ultimo inserito) " + vbCrLF + _
      "     SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(LTRIM(RTRIM(str_log_codAlb)), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
      "         FROM tb_str_logs " + vbCrLF + _
      "         WHERE LTRIM(RTRIM(str_log_codAlb)) LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "               LEN(LTRIM(RTRIM(str_log_codAlb))) = @CODE_LENGHT " + vbCrLF + _
      vbCrLF + _
	  "		--recupera dalle strutture la parte incrementale nel caso in cui nel log non ci sia " + vbCrLF + _
	  "     IF @VAR_CODEPART = 0 " + vbCrLF + _
	  "     	SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(LTRIM(RTRIM(CodAlb)), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
	  "     		FROM tb_loginstru " + vbCrLF + _
	  "     		WHERE LTRIM(RTRIM(CodAlb)) LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "     			  LEN(LTRIM(RTRIM(CodAlb))) = @CODE_LENGHT " + vbCrLF + _
	  vbCrLF + _  
      "     --incrementa parte variabile " + vbCrLF + _
      "     SET @VAR_CODEPART = @VAR_CODEPART + 1 " + vbCrLF + _
      vbCrLF + _
      "     --compone codice regionale definitivo " + vbCrLF + _
      "     SET @REGCODE = @REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART AS NVARCHAR(12)) " + vbCrLF + _
      vbCrLF + _
	  "     RETURN @REGCODE " + vbCrLF + _
	  " END " + vbCrLF
CALL DB.Execute(sql, 1295)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1296
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__162(conn)
CALL DB.Execute(sql, 1296)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__162(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1297
'...........................................................................................
'	20/09/2010 - Giacomo
'...........................................................................................
'	copia tutti gli alberghi con servizio congressuale nella sezione alberghi turismo congressuale
'...........................................................................................
sql = " DECLARE turismo_congressuale CURSOR " + vbCrLF + _
	  " READ_ONLY " + vbCrLF + _
	  " FOR SELECT RegCode, Denominazione, Comune  " + vbCrLF + _
	  " 	FROM VIEW_Strutture " + vbCrLF + _
	  " 	WHERE Modello=18  " + vbCrLF + _
	  " 		AND regCode NOT IN (SELECT CodAlb_Replicato FROM tb_loginStru WHERE Modello=40) " + vbCrLF + _
	  " 		AND (SELECT Count(*) FROM rel_str_serv WHERE rel_id_str_serv=Str_ID AND rel_str_id_relserv IN  " + vbCrLF + _
	  " 		(SELECT rel_grp_serv_id FROM rel_grp_serv WHERE rel_grp_id_serv IN (123)))=1 " + vbCrLF + _
	  "  " + vbCrLF + _
	  " DECLARE @str_id int " + vbCrLF + _
	  " DECLARE @regcode varchar(13) " + vbCrLF + _
	  " DECLARE @denominazione varchar(255) " + vbCrLF + _
	  " DECLARE @comune varchar(6) " + vbCrLF + _
	  " DECLARE @modello int " + vbCrLF + _
	  " DECLARE @tipo int " + vbCrLF + _
	  " DECLARE @utente int " + vbCrLF + _
	  " DECLARE @login varchar(10) " + vbCrLF + _
	  " DECLARE @operazione varchar(255) " + vbCrLF + _
	  "  " + vbCrLF + _
	  " SET @modello = 40 " + vbCrLF + _
	  " SET @tipo = (SELECT TOP 1 Tip_ID FROM tb_tipi_str WHERE Tip_Mod_Id = @modello) " + vbCrLF + _
	  " SET @utente = 1 " + vbCrLF + _
	  " SET @login = 'NEXTAIM' " + vbCrLF + _
	  " SET @operazione = 'Scheda principale: Inserimento nuova struttura.' " + vbCrLF + _
	  " OPEN turismo_congressuale " + vbCrLF + _
	  "   " + vbCrLF + _
	  " FETCH NEXT FROM turismo_congressuale INTO @regcode, @denominazione, @comune " + vbCrLF + _
	  "  " + vbCrLF + _
	  " WHILE (@@fetch_status <> -1) " + vbCrLF + _
	  "  " + vbCrLF + _
	  " BEGIN " + vbCrLF + _
	  " 	IF (@@fetch_status <> -2) " + vbCrLF + _
	  " 		BEGIN " + vbCrLF + _
	  " 			EXEC spstr_INSERT_NEW @denominazione, @comune, @modello, @tipo, @utente, @str_id OUTPUT " + vbCrLF + _
	  " 			UPDATE tb_loginStru SET CodAlb_Replicato = @regcode WHERE CODALB LIKE (SELECT CODALB FROM VIEW_valid_strutture WHERE Str_ID = @str_id) " + vbCrLF + _
	  " 			EXEC spstr_WRITE_LOG @str_id, NULL, @login, @operazione, 1 " + vbCrLF + _
	  " 		END " + vbCrLF + _
	  " 	FETCH NEXT FROM turismo_congressuale INTO @regcode, @denominazione, @comune " + vbCrLF + _
	  " END " + vbCrLF + _
	  "  " + vbCrLF + _
	  " CLOSE turismo_congressuale " + vbCrLF + _
	  " DEALLOCATE turismo_congressuale "
CALL DB.Execute(sql, 1297)
if DB.last_update_executed then
	dim rs_temp, rs_up
	set rs_temp = Server.CreateObject("ADODB.RecordSet")
	set rs_up = Server.CreateObject("ADODB.RecordSet")
	sql = " SELECT CodAlb FROM tb_loginstru WHERE Modello=18 " + _
		  " AND (SELECT Count(*) FROM rel_str_serv WHERE rel_id_str_serv=current_valid_str_id AND rel_str_id_relserv IN " + _
		  " (SELECT rel_grp_serv_id FROM rel_grp_serv WHERE rel_grp_id_serv IN (123)))=1 "
	rs_temp.open sql, conn, adOpenDynamic, adLockOptimistic
	while not rs_temp.eof
		CALL AggiornaStrutturaReplicata(conn, rs_up, rs_temp("CodAlb"), "")
		rs_temp.moveNext
	wend
	rs_temp.close
	set rs_temp = nothing
	set rs_up = nothing
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1298
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__163(conn)
CALL DB.Execute(sql, 1298)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1299
'...........................................................................................
'	28/09/2010 - Giacomo
'...........................................................................................
'	aggiunge i campi usati sulla sezione "anagrafica" degli agriturismi
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		Registro_numero nvarchar (60) NULL , " + _
	  "		Registro_data smalldatetime NULL;"
CALL DB.Execute(sql, 1299)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1299)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1300
'...........................................................................................
'	28/09/2010 - Giacomo
'...........................................................................................
'	aggiunge campo per etichetta nell'associazione tra dotazione e raggruppamento
'...........................................................................................
sql = " ALTER TABLE rel_grp_dotaz ADD " + _
	  "		rel_etichetta_it nvarchar(255) NULL ; "
CALL DB.Execute(sql, 1300)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1300)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1301
' Sergio 29/09/2010 - aggiunge colonne mancanti a 
' tb_contents_index 
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__164(conn)
CALL DB.Execute(sql, 1301)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD_EX(1301,"v_indice")
CALL DB.SqlServer_VIEWS_REBUILD_EX(1301,"v_indice_visibile")
CALL DB.SqlServer_VIEWS_REBUILD(1301)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1302
'...........................................................................................
'	14/10/2010 - Giacomo 
'...........................................................................................
'	ricrea la stored procedure e la ricrea per apportare una correzione
'...........................................................................................
sql = DropObject(conn, "spstr_SET_APT_LIST", "PROCEDURE") + _
	  " CREATE PROCEDURE [dbo].[spstr_SET_APT_LIST]( " + vbCrLF + _
	  " @RegCode nvarchar(12), " + vbCrLF + _
	  " @TYPE nvarchar(2) " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " DECLARE @APTCode nvarchar(50) " + vbCrLF + _
	  " IF (@TYPE='T') BEGIN		--gestione tipologie " + vbCrLF + _
	  " 		DECLARE @Cod_Proprietario nvarchar(12) " + vbCrLF + _
	  " " + vbCrLF + _
	  " 		SELECT TOP 1 @Cod_Proprietario=Cod_Proprietario FROM tb_Strutture WHERE RegCode=@RegCode ORDER BY DataModifica DESC " + vbCrLF + _
	  " 		SELECT TOP 1 @APTCODE = AptCode FROM tb_strutture WHERE RegCode=@Cod_Proprietario ORDER BY DataModifica DESC " + vbCrLF + _
	  " 		UPDATE tb_strutture SET AptCode=@APTCODE WHERE tb_strutture.RegCode=@RegCode " + vbCrLF + _
	  " 	END " + vbCrLF + _
	  " ELSE BEGIN		--gestione proprietari / agenzie " + vbCrLF + _
	  " 	DECLARE @APT_LIST nvarchar(50) " + vbCrLF + _
	  " 	SET @APT_LIST = '' " + vbCrLF + _
	  " 	DECLARE rsapt CURSOR FOR " + vbCrLF + _
	  " 			SELECT DISTINCT AptCode FROM tb_Strutture WHERE Cod_Proprietario=@RegCode " + vbCrLF + _
	  " 	OPEN rsapt " + vbCrLF + _
	  " 	FETCH NEXT FROM rsapt INTO @APTCODE " + vbCrLF + _
	  " 	WHILE (@@FETCH_STATUS=0) " + vbCrLF + _
	  " 		BEGIN " + vbCrLF + _
	  " 			SET @APT_LIST = REPLACE(@APT_LIST,' ' + @APTCODE + ' ','') " + vbCrLF + _
	  " 			SET @APT_LIST = RTRIM(LTRIM(@APT_LIST)) + ' ' + RTRIM(LTRIM(@APTCODE)) + ' ' " + vbCrLF + _
	  " 			FETCH NEXT FROM rsapt INTO @APTCODE " + vbCrLF + _
	  " 		END " + vbCrLF + _
	  " 	CLOSE rsapt " + vbCrLF + _
	  " 	DEALLOCATE rsapt " + vbCrLF + _
	  " 	--imposta la lista di apt per il record corrente " + vbCrLF + _
	  " 	IF RTRIM(@APT_LIST)<>'' " + vbCrLF + _
	  " 		UPDATE tb_strutture SET AptCode = @APT_LIST WHERE RegCode=@REGCODE " + vbCrLF + _
	  " END "
CALL DB.Execute(sql, 1302)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1303
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__165(conn)
CALL DB.Execute(sql, 1303)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1304
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__166(conn)
CALL DB.Execute(sql, 1304)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1305
'...........................................................................................
'	03/11/2010 - Giacomo
'...........................................................................................
'	aggiunge campo per testo di presentazione della struttura / azienda
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD " + _
	  "		presentazione_struttura ntext NULL ; "
CALL DB.Execute(sql, 1305)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1306
'...........................................................................................
'	15/11/2010 - Nicola
'...........................................................................................
'	Imposta testi di default per la dichiarazione
'...........................................................................................
sql = " UPDATE tb_modelli SET " + _
		" mod_intestazione_premessa = '', " + _
		" mod_intestazione_label_salva = 'Consente di salvare temporaneamente le modifiche e poter in seguito procedere ad ulteriori aggiornamenti online.', " + _
		" mod_intestazione_label_anteprima = 'Consente di  generare una stampa di prova con gli ultimi dati salvati. ', " + _
		" mod_intestazione_label_presenta = 'Funzione da utilizzare quando si è completato l''inserimento dei dati.<br>" + vbCrLf + _
										    "Premendo questo pulsante si presenta in forma telematica la dichiarazione, inviando la segnalazione all''ufficio competente.', " + _
		" mod_intestazione_label_chiudi = 'Consente di uscire da questa sezione lasciando inalterati i dati precedentemente salvati. ', " + _
		" mod_intestazione_istruzioni = 'Per maggiori informazioni, per problemi o difficoltà nella compilazione della dichiarazione o per segnalazioni riguardanti i dati presenti contattare:" + vbCrLF + _
										"APT della Provincia di Venezia" + vbCrLf + _
										"Castello 5050, 30122 Venezia" + vbCrLf + _
										"Tel. 041.5298711 fax 041.5230399" + vbCrLf + _
										"E-mail: info@turismovenezia.it ' " + _
	  " WHERE mod_tipo_Record LIKE 'K' OR mod_tipo_Record LIKE 'C' "
CALL DB.Execute(sql, 1306)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1307
'...........................................................................................
'	16/11/2010 - Nicola
'...........................................................................................
'	Imposta testi ed email per la dichiarazione
'...........................................................................................
sql = " UPDATE tb_modelli SET " + _
		" mod_email_sender = 'sviluppo@combinario.com', " + _
		" mod_email_validazione_oggetto = 'Apt di Venezia - Convention Bureau - dichiarazione completata correttamente', " + _
		" mod_email_conferma_oggetto = 'Apt di Venezia - Convention Bureau - dichiarazione presentata correttamente', " + _
		" mod_email_password_oggetto = 'Apt di Venezia - costituzione del Convention Bureau', " + _
		" mod_page_email_conferma = 1592, " + _
		" mod_page_email_validazione = 1602, " + _
		" mod_page_email_password = 1581 " + _
	  " WHERE mod_tipo_Record LIKE 'K' OR mod_tipo_Record LIKE 'C' "
CALL DB.Execute(sql, 1307)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1308
'...........................................................................................
'	18/11/2010 - Giacomo
'...........................................................................................
'	Crea trigger che mantengono sincronizzati i dati di login, password e stato abilitazione
'	tra tb_utenti e tb_loginStru;
'	Update che sincronizza tb_utenti e tb_loginStru;
'...........................................................................................
sql = " CREATE TRIGGER [tb_loginStru_update] " + vbCrLf + _
	  " ON [dbo].[tb_loginStru]  " + vbCrLf + _
	  " AFTER UPDATE  " + vbCrLf + _
	  " AS  " + vbCrLf + _
	  "  " + vbCrLf + _
	  " DECLARE @counter int " + vbCrLf + _
	  " SELECT @counter = COUNT(*) FROM (tb_Utenti u INNER JOIN tb_Indirizzario i ON u.ut_NextCom_ID = i.IDElencoIndirizzi)  " + vbCrLf + _
	  " 					INNER JOIN inserted ON RTRIM(LTRIM(i.SyncroKey)) = RTRIM(LTRIM(inserted.CODALB)) " + vbCrLf + _
	  " 					WHERE SyncroTable LIKE '%VIEW_valid_strutture%' " + vbCrLf + _
	  " 					AND (ut_login <> inserted.[LOGIN] OR ut_password <> inserted.[PASSWORD] OR ut_Abilitato <> inserted.Struttura_attiva) " + vbCrLf + _
	  " IF (@counter >0) " + vbCrLf + _
	  " BEGIN " + vbCrLf + _
	  " 	UPDATE tb_Utenti SET ut_login = inserted.LOGIN, ut_password = inserted.PASSWORD, ut_Abilitato = inserted.Struttura_attiva  " + vbCrLf + _
	  " 	FROM (tb_Utenti u INNER JOIN tb_Indirizzario i ON u.ut_NextCom_ID = i.IDElencoIndirizzi)  " + vbCrLf + _
	  " 	INNER JOIN inserted ON RTRIM(LTRIM(i.SyncroKey)) = RTRIM(LTRIM(inserted.CODALB)) " + vbCrLf + _
	  " 	WHERE SyncroTable LIKE '%VIEW_valid_strutture%' " + vbCrLf + _
	  " END " + vbCrLf + _
	  " ; " + vbCrLf + _
	  " CREATE TRIGGER [tb_Utenti_update]  " + vbCrLf + _
	  " ON [dbo].[tb_Utenti]  " + vbCrLf + _
	  " AFTER UPDATE  " + vbCrLf + _
	  " AS " + vbCrLf + _
	  "  " + vbCrLf + _
	  " DECLARE @counter int " + vbCrLf + _
	  " SELECT @counter = COUNT(*) FROM (tb_loginStru l INNER JOIN tb_Indirizzario i ON  RTRIM(LTRIM(l.CODALB)) = RTRIM(LTRIM(i.SyncroKey))) " + vbCrLf + _
	  " 					INNER JOIN inserted ON i.IDElencoIndirizzi = inserted.ut_NextCom_ID " + vbCrLf + _
	  " 					WHERE i.SyncroTable LIKE '%VIEW_valid_strutture%' " + vbCrLf + _
	  " 					AND (l.[LOGIN] <> inserted.ut_login OR l.[PASSWORD] <> inserted.ut_password OR l.Struttura_attiva <> inserted.ut_Abilitato) " + vbCrLf + _
	  "  " + vbCrLf + _
	  " IF (@counter > 0) " + vbCrLf + _
	  " BEGIN " + vbCrLf + _
	  " 	UPDATE tb_loginStru SET [LOGIN] = inserted.ut_login, [PASSWORD] = inserted.ut_password, Struttura_attiva = inserted.ut_Abilitato " + vbCrLf + _
	  " 	FROM (tb_loginStru l INNER JOIN tb_Indirizzario i ON  RTRIM(LTRIM(l.CODALB)) = RTRIM(LTRIM(i.SyncroKey))) " + vbCrLf + _
	  " 	INNER JOIN inserted ON i.IDElencoIndirizzi = inserted.ut_NextCom_ID " + vbCrLf + _
	  " 	WHERE i.SyncroTable LIKE '%VIEW_valid_strutture%' " + vbCrLf + _
	  " END " + vbCrLf + _
	  " ; " + vbCrLf + _
	  " UPDATE tb_Utenti SET ut_login = tb_loginStru.[LOGIN], ut_password = tb_loginStru.[PASSWORD], ut_Abilitato = tb_loginStru.Struttura_attiva " + vbCrLf + _
	  " FROM (tb_Utenti u INNER JOIN tb_Indirizzario i ON u.ut_NextCom_ID = i.IDElencoIndirizzi) " + vbCrLf + _
	  " INNER JOIN tb_loginStru ON RTRIM(LTRIM(i.SyncroKey)) = RTRIM(LTRIM(tb_loginStru.CODALB)) " + vbCrLf + _
	  " WHERE SyncroTable LIKE '%VIEW_valid_strutture%' " + vbCrLf + _
	  " AND (ut_login <> tb_loginStru.[LOGIN] OR ut_password <> tb_loginStru.[PASSWORD] OR ut_Abilitato <> tb_loginStru.Struttura_attiva) "
CALL DB.Execute(sql, 1308)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1309
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__167(conn) + _
	  Aggiornamento__FRAMEWORK_CORE__168(conn)
CALL DB.Execute(sql, 1309)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1310
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__169(conn)
CALL DB.Execute(sql, 1310)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1310)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1311
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__170(conn)
CALL DB.Execute(sql, 1311)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1311)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1312
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__171(conn)
CALL DB.Execute(sql, 1312)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1312)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1313
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__172(conn)
CALL DB.Execute(sql, 1313)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1313)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1314
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__173(conn)
CALL DB.Execute(sql, 1314)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1314)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1315
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__174(conn)
CALL DB.Execute(sql, 1315)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1315)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1316
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__175(conn)
CALL DB.Execute(sql, 1316)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__175(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1317
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__176(conn)
CALL DB.Execute(sql, 1317)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1317)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1318
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__177(conn)
CALL DB.Execute(sql, 1318)
'*******************************************************************************************


'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1318)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1319
'...........................................................................................
'	07/12/2010 - Nicola
'...........................................................................................
'	corregge dati per export regione
'...........................................................................................
sql = " UPDATE rel_grp_dotaz SET " + _
	  "		rel_sez_xml_rvtweb = 'PREZZI' " + _
	  " WHERE rel_grp_id_dotaz IN (220, 221,222, 223,358) "
CALL DB.Execute(sql, 1319)
'*******************************************************************************************



'*******************************************************************************************
' <-- qua AGGIORNAMENTI U.A.
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO 1320
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__178(conn)
CALL DB.Execute(sql, 1320)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1320)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1321
'...........................................................................................
'aggiunge campo su tabella strutture per Partita Iva
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD Lic_PIva " & SQL_CharField(conn,60) & "  null; "
CALL DB.Execute(sql, 1321)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1322
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__179(conn)
CALL DB.Execute(sql, 1322)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1323
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__180(conn)
CALL DB.Execute(sql, 1323)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1324
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__181(conn)
CALL DB.Execute(sql, 1324)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1325
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__182(conn)
CALL DB.Execute(sql, 1325)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1326
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__183(conn)
CALL DB.Execute(sql, 1326)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1327
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__184(conn)
CALL DB.Execute(sql, 1327)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1328
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__185(conn)
CALL DB.Execute(sql, 1328)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1329
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__186(conn)
CALL DB.Execute(sql, 1329)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1330
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__187(conn)
CALL DB.Execute(sql, 1330)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1331
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__188(conn)
CALL DB.Execute(sql, 1331)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1332
'...........................................................................................
'	05/04/2011 - Giacomo
'...........................................................................................
'	modifica stored procedure per il calcolo del codice regionale
'...........................................................................................
sql = DropObject(conn, "fn_NEW_REGCODE", "FUNCTION") + _
	  " CREATE FUNCTION dbo.fn_NEW_REGCODE(" + vbCrLF + _
	  "		@COMUNE nvarchar(6), " + vbCrLF + _
      "		@MODELLO int " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " RETURNS nvarchar(12) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " BEGIN " + vbCrLF + _
	  "     DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
	  "     DECLARE @FIRST_LETTER nvarchar(3) " + vbCrLF + _
	  "     DECLARE @VAR_CODEPART int " + vbCrLF + _
	  "     DECLARE @CODE_LENGHT int " + vbCrLF + _
      "     DECLARE @VAR_CODEPART_LENGTH int " + vbCrLF + _
      "     DECLARE @CHECK_EXIST int " + vbCrLF + _
      vbCrLF + _
      "     --recupera radicie del codice regionale " + vbCrLF + _
	  "     SELECT @FIRST_LETTER = Mod_FirstLT_Regcode FROM tb_Modelli WHERE Mod_ID= @MODELLO " + vbCrLF + _
      vbCrLF + _
      "     --compone prima parte del codice regionale " + vbCrLF + _
      "     SET @REGCODE = @FIRST_LETTER + @COMUNE " + vbCrLF + _
      vbCrLF + _
      "     --calcola lunghezza parte incrementale del codice " + vbCrLF + _
      "     IF (LEN(@FIRST_LETTER)) > 1 " + vbCrLF + _
      "         SET @CODE_LENGHT = 12 " + vbCrLF + _
      "     ELSE " + vbCrLF + _
	  "         SET @CODE_LENGHT = 11 " + vbCrLF + _
	  "     SET @VAR_CODEPART_LENGTH = @CODE_LENGHT - LEN(@REGCODE) " + vbCrLF + _
      vbCrLF + _
      "     --recupera parte incrementale del codice (ultimo inserito) " + vbCrLF + _
      "     SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(LTRIM(RTRIM(str_log_codAlb)), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
      "         FROM tb_str_logs " + vbCrLF + _
      "         WHERE LTRIM(RTRIM(str_log_codAlb)) LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "               LEN(LTRIM(RTRIM(str_log_codAlb))) = @CODE_LENGHT " + vbCrLF + _
	  vbCrLF + _
	  "		--controllo se esiste già una struttura con il regcode che verrà utilizzato. " + vbCrLF + _
	  "		--serve per non dare errore nel caso in cui non ci sia nessun record nel log (str_log_codAlb) associato all'ultima struttura inserita " + vbCrLF + _
	  "		SELECT  @CHECK_EXIST = CAST(ISNULL(current_str_id, 0) AS INT) FROM tb_loginstru  " + vbCrLF + _
	  "			WHERE CODALB LIKE  " + vbCrLF + _
	  "			@REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART + 1 AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART + 1 AS NVARCHAR(12)) " + vbCrLF + _
      vbCrLF + _
	  "		--recupera dalle strutture la parte incrementale nel caso in cui nel log non ci sia " + vbCrLF + _
	  "     IF @VAR_CODEPART = 0 OR @CHECK_EXIST > 0 " + vbCrLF + _
	  "     	SELECT @VAR_CODEPART = ISNULL(CAST(MAX(RIGHT(LTRIM(RTRIM(CodAlb)), @VAR_CODEPART_LENGTH)) AS Int),0) " + vbCrLF + _
	  "     		FROM tb_loginstru " + vbCrLF + _
	  "     		WHERE LTRIM(RTRIM(CodAlb)) LIKE (@REGCODE + '%') AND " + vbCrLF + _
	  "     			  LEN(LTRIM(RTRIM(CodAlb))) = @CODE_LENGHT " + vbCrLF + _
	  vbCrLF + _  
      "     --incrementa parte variabile " + vbCrLF + _
      "     SET @VAR_CODEPART = @VAR_CODEPART + 1 " + vbCrLF + _
      vbCrLF + _
      "     --compone codice regionale definitivo " + vbCrLF + _
      "     SET @REGCODE = @REGCODE + REPLICATE(0, (@VAR_CODEPART_LENGTH - LEN(CAST(@VAR_CODEPART AS NVARCHAR(12))) )) + CAST(@VAR_CODEPART AS NVARCHAR(12)) " + vbCrLF + _
      vbCrLF + _
	  "     RETURN @REGCODE " + vbCrLF + _
	  " END " + vbCrLF
CALL DB.Execute(sql, 1332)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1333
'corregge comune di una struttura ricettiva (u.a.n.c. Saltarel ada)
'Aggiunge dati per correzione stored procedure apt list
'Nicola  12/04/2011
'...........................................................................................
sql = "UPDATE tb_strutture SET comune = '027044' WHERE RegCode LIKE 'M0270420245'; " + _
	  DropObject(conn, "spstr_SET_APT_LIST", "PROCEDURE") + _
	  " CREATE PROCEDURE [dbo].[spstr_SET_APT_LIST]( " + vbCrLF + _
	  " @RegCode nvarchar(12), " + vbCrLF + _
	  " @TYPE nvarchar(2) " + vbCrLF + _
	  " ) " + vbCrLF + _
	  " AS " + vbCrLF + _
	  " DECLARE @APTCode nvarchar(50) " + vbCrLF + _
	  " IF (@TYPE='T') BEGIN		--gestione tipologie " + vbCrLF + _
	  " 		DECLARE @Cod_Proprietario nvarchar(12) " + vbCrLF + _
	  " " + vbCrLF + _
	  " 		SELECT TOP 1 @Cod_Proprietario=Cod_Proprietario FROM tb_Strutture WHERE RegCode=@RegCode ORDER BY DataModifica DESC " + vbCrLF + _
	  " 		SELECT TOP 1 @APTCODE = AptCode FROM tb_strutture WHERE RegCode=@Cod_Proprietario ORDER BY DataModifica DESC " + vbCrLF + _
	  " 		UPDATE tb_strutture SET AptCode=@APTCODE WHERE tb_strutture.RegCode=@RegCode " + vbCrLF + _
	  " 	END " + vbCrLF + _
	  " ELSE BEGIN		--gestione proprietari / agenzie " + vbCrLF + _
	  " 	DECLARE @APT_LIST nvarchar(50) " + vbCrLF + _
	  " 	SET @APT_LIST = '' " + vbCrLF + _
	  " 	DECLARE rsapt CURSOR FOR " + vbCrLF + _
	  " 			SELECT DISTINCT AptCode FROM view_records_Strutture WHERE Cod_Proprietario=@RegCode AND mod_tipo_record <> 'T' " + vbCrLF + _
	  " 	OPEN rsapt " + vbCrLF + _
	  " 	FETCH NEXT FROM rsapt INTO @APTCODE " + vbCrLF + _
	  " 	WHILE (@@FETCH_STATUS=0) " + vbCrLF + _
	  " 		BEGIN " + vbCrLF + _
	  " 			SET @APT_LIST = REPLACE(@APT_LIST,' ' + @APTCODE + ' ','') " + vbCrLF + _
	  " 			SET @APT_LIST = RTRIM(LTRIM(@APT_LIST)) + ' ' + RTRIM(LTRIM(@APTCODE)) + ' ' " + vbCrLF + _
	  " 			FETCH NEXT FROM rsapt INTO @APTCODE " + vbCrLF + _
	  " 		END " + vbCrLF + _
	  " 	CLOSE rsapt " + vbCrLF + _
	  " 	DEALLOCATE rsapt " + vbCrLF + _
	  " 	--imposta la lista di apt per il record corrente " + vbCrLF + _
	  " 	IF RTRIM(@APT_LIST)<>'' " + vbCrLF + _
	  " 		UPDATE tb_strutture SET AptCode = @APT_LIST WHERE RegCode=@REGCODE " + vbCrLF + _
	  " END "
CALL DB.Execute(sql, 1333)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1333(DB.objConn, rs, rst, "<>'T'")
	CALL AggiornamentoSpeciale_1333(DB.objConn, rs, rst, "='T'")
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub AggiornamentoSpeciale_1333(conn, rs, rst, mod_tipo_record_condition)
	dim sql, readConn, readRs
	sql = " SELECT str_id FROM VIEW_Strutture WHERE aptcode like '%06%' AND aptcode LIKE '%15%' AND mod_tipo_record" & mod_tipo_record_condition
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		CALL SET_APT(conn, rst, rs("str_id"))
		rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1334
'Aggiunge dati per attivazione ambito Caorle Bis
'Nicola  13/04/2011
'...........................................................................................
sql = " INSERT INTO tb_apt (apt_codice, apt_nome) VALUES ('16', 'N.4 Bis Caorle') ; " + _
	  " UPDATE tb_apt SET apt_nome = 'N.4 Bibione' WHERE apt_codice LIKE '04'; " + _
	  " UPDATE tb_comuni SET cod_APT = '04', ufficio_apt = 2 WHERE Codice_ISTAT LIKE '027016' OR Codice_ISTAT LIKE '027040'; " + _
	  " UPDATE tb_apt_uffici SET uf_nome = 'N.4 Bis Caorle', uf_apt_codice = '16' WHERE uf_id=1; " + _
	  " UPDATE tb_apt_uffici SET uf_nome = 'N.4 Bibione', uf_apt_codice = '04' WHERE uf_id = 2 ; " + _
	  " UPDATE tb_comuni SET cod_APT = '16' WHERE ufficio_apt=1; "
CALL DB.Execute(sql, 1334)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1334(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub AggiornamentoSpeciale_1334(conn, rs, rst)
	dim sql, readConn, readRs
	sql = " SELECT RegCode, mod_tipo_record, str_id FROM VIEW_Strutture " + _
		  " WHERE mod_tipo_record <> 'P' AND mod_tipo_record <> 'C' AND mod_tipo_record <> 'K' AND mod_tipo_record <> 'R' " + _
				" AND AptCode LIKE '%04%' " + _
		  " ORDER BY ( CASE WHEN mod_tipo_record LIKE 'S'  THEN 0 " + _
					 " WHEN mod_tipo_record LIKE 'U' THEN 1 " + _
					 " WHEN mod_tipo_record LIKE 'A' THEN 2 " + _
					 " WHEN mod_tipo_record LIKE 'O' THEN 3 " + _
					 " WHEN mod_tipo_record LIKE 'T' THEN 4 END) "
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		CALL SET_APT(conn, rst, rs("str_id"))
		rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1335
'Correzzion dati apt e divisione bibione e caorle
'Nicola  13/04/2011
'...........................................................................................
sql = " update tb_strutture " + _
	  " set aptCode = cod_apt " + _
	  " From tb_strutture inner join tb_comuni on tb_strutture.comune = tb_comuni.codice_istat " + _
	  " inner join tb_tipi_str on tb_Strutture.tipo = tb_tipi_str.tip_id " + _
	  " inner join tb_modelli on tb_tipi_str.tip_mod_id = tb_modelli.mod_id " + _
	  " where ufficio_apt IN (1,2) and mod_tipo_record like 'S' "
CALL DB.Execute(sql, 1335)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1334(DB.objConn, rs, rst)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1336
'Corregge associazioni unità abitative di caorle ad agenzie di San Michele al Tagliamento
'Nicola  19/04/2011
'...........................................................................................
sql = " UPDATE tb_strutture SET Cod_Proprietario='IP0270050003' WHERE str_id=132611 ; " + _
      " UPDATE tb_strutture SET Cod_Proprietario='IP0270050040' WHERE str_id=106753 ; " + _
	  " UPDATE tb_strutture SET Cod_Proprietario='IP0270050040' WHERE str_id=106758 ; "
CALL DB.Execute(sql, 1336)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale_1336(DB.objConn, rs, rst)
end if
'...........................................................................................
'funzione dichiarata per evitare interferenze con altre variabili d'ambiente
sub AggiornamentoSpeciale_1336(conn, rs, rst)
	dim sql, readConn, readRs
	sql = " SELECT RegCode, mod_tipo_record, str_id FROM VIEW_Strutture " + _
		  " WHERE mod_tipo_record <> 'P' AND mod_tipo_record <> 'C' AND mod_tipo_record <> 'K' AND mod_tipo_record <> 'R' " + _
				" AND AptCode LIKE '%04%' AND AptCode LIKE '%16%' " + _
		  " ORDER BY ( CASE WHEN mod_tipo_record LIKE 'S'  THEN 0 " + _
					 " WHEN mod_tipo_record LIKE 'U' THEN 1 " + _
					 " WHEN mod_tipo_record LIKE 'A' THEN 2 " + _
					 " WHEN mod_tipo_record LIKE 'O' THEN 3 " + _
					 " WHEN mod_tipo_record LIKE 'T' THEN 4 END) "
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		CALL SET_APT(conn, rst, rs("str_id"))
		rs.movenext
	wend
	rs.close
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1337
'...........................................................................................
'	19/04/2011 - Nicola
'...........................................................................................
'	imposta valori per data inizio attività su alberghi
'...........................................................................................
sql = " UPDATE tb_strutture " + _
	  " SET inizio_attivita = DecretoCL_D " + _
	  " WHERE str_id IN (SELECT str_id FROM view_testata_Strutture WHERE modello=18) "
CALL DB.Execute(sql, 1337)
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 1338
'cancella gli agriturismi che hannno Lic_Numero NULL
'Giacomo  19/04/2011
'...........................................................................................
sql = " SELECT Lic_numero, REGCODE FROM VIEW_strutture WHERE modello = 52 " & _
	  " AND ISNULL(Lic_numero, '') = '' "
CALL DB.Execute(sql, 1338)
if DB.last_update_executed then
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		'registrazione su log
		CALL WRITE_log(conn, rs("REGCODE"), NULL, Session("LOGIN_4_LOG"), "Cancellazione completa struttura", false)
		
		'cancella contatto sincronizzato dal NEXT-com e NEXT-INFO
		CALL CancellaStruttura_NextCom_NextInfo(conn, rs("REGCODE"))
		
		'cancellazione intera struttura
		sql = "EXEC DELETE_Struttura '" & rs("REGCODE") & "'"
		CALL conn.Execute(sql,0,adCmdText)
		
		rs.movenext
	wend
	rs.close
end if
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO 1339
' Riporta i valori delle vecchie dotazioni su quelle nuove per gli agriturismi
'Giacomo  19/04/2011
'...........................................................................................
sql = " SELECT MIN(str_id) AS str_id ,RegCode FROM VIEW_records_strutture WHERE Mod_id = 52 GROUP BY RegCode "
CALL DB.Execute(sql, 1339)
if DB.last_update_executed then
	dim ins_str_id, i, arr(15), tipo
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	while not rs.eof
		sql = "SELECT MAX(str_id) FROM VIEW_records_strutture WHERE Mod_id = 52 AND regcode LIKE '"&rs("RegCode")& "'"
		ins_str_id = GetValueList(conn,NULL,sql)

		'COLTURE - (gruppo_id = 297, dotaz_id = 520)
		sql = "SELECT rel_grp_dotaz_id FROM rel_grp_dotaz WHERE rel_id_Grp_dotaz = 297 AND rel_grp_id_dotaz = 520 ORDER BY rel_grp_dotaz_id"
		rst.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		i = 1
		while not rst.eof
			if i<16 then
				arr(i) = rst("rel_grp_dotaz_id")
			else
				response.write "aumentare numero colture, i="&i&" <br>"
			end if
			rst.moveNext
			i = i + 1
		wend
		rst.close
		
		'recupero i valori delle vecchie dotazioni del gruppo colture - pos 1 (superficie)
		sql = " SELECT dotaz_nome_it, rel_str_dotaz_valore FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz " & _
			  " ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id INNER JOIN tb_dotazioni " & _
			  " ON rel_Grp_dotaz.rel_Grp_id_dotaz = tb_dotazioni.dotaz_id " & _
			  " WHERE rel_id_Grp_dotaz = 297 AND rel_Grp_id_dotaz IN (550,553,556,559,562,565,568,571,574,577,580,583,586,589,592,595,598,601,604,607,610,613,616,619,622,625,628) " & _
			  " AND rel_str_dotaz_pos_val = 1 AND rel_id_str_dotaz = " & rs("str_id") & _
			  " ORDER BY rel_str_dotaz_id "
		rst.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		i = 1
		'inserisco valori su nuove dotazioni
		while not rst.eof
			'inserisco pos 1 - tipo
			tipo = rst("dotaz_nome_it")
			tipo = Replace(tipo,"(superficie)","")
			tipo = Trim(tipo)
			sql = " INSERT INTO rel_str_dotaz(rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val)" & _
				  " VALUES ("&ins_str_id&","&arr(i)&",'"&tipo&"', 1) "
			CALL conn.Execute(sql)
				   
			'inserisco pos 2 - superficie
			sql = " INSERT INTO rel_str_dotaz(rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val)" & _
				  " VALUES ("&ins_str_id&","&arr(i)&",'"&rst("rel_str_dotaz_valore")&"', 2) "
			CALL conn.Execute(sql)
			
			rst.moveNext
			i = i + 1
		wend
		rst.close
		

	
		'ALLEVAMENTI - (gruppo_id = 298, dotaz_id = 521)
		sql = " SELECT rel_grp_dotaz_id FROM rel_grp_dotaz WHERE rel_id_Grp_dotaz = 298 AND rel_grp_id_dotaz = 521 ORDER BY rel_grp_dotaz_id "
		rst.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		i = 1
		while not rst.eof
			if i<16 then
				arr(i) = rst("rel_grp_dotaz_id")
			else
				response.write "aumentare numero allevamenti, i="&i&" <br>"
			end if
			rst.moveNext
			i = i + 1
		wend
		rst.close
		'recupero i valori delle vecchie dotazioni del gruppo allevamenti - pos 1 (capi presenti)
		sql = " SELECT dotaz_nome_it, rel_str_dotaz_valore FROM rel_str_dotaz INNER JOIN rel_Grp_dotaz " & _
			  " ON rel_str_dotaz.rel_str_id_dotaz = rel_Grp_dotaz.rel_Grp_dotaz_id INNER JOIN tb_dotazioni " & _
			  " ON rel_Grp_dotaz.rel_Grp_id_dotaz = tb_dotazioni.dotaz_id " & _
			  " WHERE rel_id_Grp_dotaz = 298 AND rel_Grp_id_dotaz IN (631,632,633,634,635,636,637,638,639,640,641,642,643,644,645,646,647,648,649,650,651,652,653,654) " & _
			  " AND rel_str_dotaz_pos_val = 2 AND rel_id_str_dotaz = " & rs("str_id") & _
			  " ORDER BY rel_str_dotaz_id "
		rst.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		i = 1
		'inserisco valori su nuove dotazioni
		while not rst.eof
			'inserisco pos 1 - tipo
			tipo = rst("dotaz_nome_it")
			tipo = Trim(tipo)
			sql = " INSERT INTO rel_str_dotaz(rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val)" & _
				  " VALUES ("&ins_str_id&","&arr(i)&",'"&tipo&"', 1) "
			CALL conn.Execute(sql)
				   
			'inserisco pos 2 - n. capi
			sql = " INSERT INTO rel_str_dotaz(rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val)" & _
				  " VALUES ("&ins_str_id&","&arr(i)&",'"&rst("rel_str_dotaz_valore")&"', 2) "
			CALL conn.Execute(sql)
			
			rst.moveNext
			i = i + 1
		wend
		rst.close
		
		ins_str_id = CHECK_AND_DUPLICATE(conn, ins_str_id)
		conn.spstr_VALIDA_RECORD ins_str_id, cString(Session("LOGIN_4_LOG")), 2010
	
		rs.movenext
	wend
	rs.close
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1340
' Aggiungo due campi a tb_stru_gest
'Giacomo  19/04/2011
'...........................................................................................
sql = " ALTER TABLE tb_stru_gest ADD azienda_dati_altro bit NULL; " & _
	  " ALTER TABLE tb_stru_gest ADD azienda_dati_note " + SQL_CharField(Conn, 0) + " NULL; "
CALL DB.Execute(sql, 1340)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1341
'...........................................................................................
sql = Aggiornamento__INFO__22(conn)
CALL DB.Execute(sql, 1341)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__INFO__22(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1342
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__189(conn)
CALL DB.Execute(sql, 1342)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1343
'...........................................................................................
sql = Aggiornamento__INFO__23(conn)
CALL DB.Execute(sql, 1343)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__INFO__23(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1344
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__190(conn)
CALL DB.Execute(sql, 1344)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1345
' Cancello delle tipologie per gli agriturismi e correggo il nome delle altre
'Giacomo  05/05/2011
'...........................................................................................
sql = " DELETE FROM tb_tipi_str WHERE Tip_Mod_id = 52 AND Tip_id IN (67,68,69,70); " & _
	  " UPDATE tb_tipi_str SET Tip_Den_it = 'Con ricettività' WHERE Tip_Mod_id = 52 AND Tip_id = 64; " & _
	  " UPDATE tb_tipi_str SET Tip_Valid_to = NULL WHERE Tip_Mod_id = 52 AND Tip_id = 64; " & _
	  " UPDATE tb_tipi_str SET Tip_Den_it = 'Senza ricettività' WHERE Tip_Mod_id = 52 AND Tip_id = 66; "
CALL DB.Execute(sql, 1345)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1346
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__191(conn)
CALL DB.Execute(sql, 1346)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1346)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1347
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__192(conn)
CALL DB.Execute(sql, 1347)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1347)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1348
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__193(conn)
CALL DB.Execute(sql, 1348)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1348)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1349
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__194(conn)
CALL DB.Execute(sql, 1349)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__194(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1350
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__195(conn)
CALL DB.Execute(sql, 1350)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1350)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1351
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__196(conn)
CALL DB.Execute(sql, 1351)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1351)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1352
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__197(conn)
CALL DB.Execute(sql, 1352)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1353
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__198(conn)
CALL DB.Execute(sql, 1353)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1354
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__199(conn)
CALL DB.Execute(sql, 1354)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1355
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__200(conn)
CALL DB.Execute(sql, 1355)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1356
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__201(conn)
CALL DB.Execute(sql, 1356)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1357
'...........................................................................................
sql = Aggiornamento__INFO__24(conn, "ru")
CALL DB.Execute(sql, 1357)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1358
'...........................................................................................
sql = Aggiornamento__INFO__24(conn, "pt")
CALL DB.Execute(sql, 1358)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1359
'...........................................................................................
sql = Aggiornamento__INFO__24(conn, "cn")
CALL DB.Execute(sql, 1359)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1360
'...........................................................................................
sql = Aggiornamento__INFO__25(conn)
CALL DB.Execute(sql, 1360)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1361
'...........................................................................................
sql = Aggiornamento__INFO__26(conn)
CALL DB.Execute(sql, 1361)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1361)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1362
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "ru")
CALL DB.Execute(sql, 1362)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1363
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "pt")
CALL DB.Execute(sql, 1363)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1364
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__202(conn, "cn")
CALL DB.Execute(sql, 1364)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1365
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__203(conn)
CALL DB.Execute(sql, 1365)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__203(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1366
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__204(conn)
CALL DB.Execute(sql, 1366)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__204(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1367
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__205(conn)
CALL DB.Execute(sql, 1367)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1368
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__206(conn)
CALL DB.Execute(sql, 1368)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1368)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1369
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__207(conn)
CALL DB.Execute(sql, 1369)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1369)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1370
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__208(conn)
CALL DB.Execute(sql, 1370)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1371
'...........................................................................................
'	aggiorna trigger alla struttura di log per il calcolo del progressivo
'	aggiunge controllo per saltare il progressivo degli agriturismi senza ricettività
'...........................................................................................
sql = DropObject(conn, "tb_str_logs_INSERT", "TRIGGER") + _
	  " CREATE TRIGGER dbo.tb_str_logs_INSERT ON tb_str_logs AFTER INSERT AS " + vbCrLf + _
      "     DECLARE @REGCODE nvarchar(13) " + vbCrLF + _
	  "     DECLARE @REGCODE_REGIONE nvarchar(13) " + vbCrLF + _
	  "		DECLARE @STR_ID nvarchar(60) " + vbCrlf + _
  	  "     DECLARE @STR_ID_REGIONE nvarchar(60) " + vbCrLF + _
      "     DECLARE @LAST INT " + vbCrLF + _
      "     DECLARE @PROGRESSIVO INT " + vbCrLF + _
      vbCrLf + _
      "     DECLARE @AGRITURISMI_CON_RIC_TIP_ID INT " + vbCrLF + _
      "     DECLARE @AGRITURISMI_MOD_ID INT " + vbCrLF + _
      "     SET @AGRITURISMI_CON_RIC_TIP_ID = 64 " + vbCrLF + _
      "     SET @AGRITURISMI_MOD_ID = 52 " + vbCrLF + _
      vbCrLF + _
	  "     SELECT @REGCODE = RTRIM(LTRIM(str_log_CodAlb)), @LAST=str_log_id, @STR_ID = str_log_record  " + vbCrLF + _
      "     FROM INSERTED WHERE ((Str_log_des LIKE '%registrazione validata%' AND  " + vbCrLF + _
      "                           (str_log_modello <> @AGRITURISMI_MOD_ID OR  " + vbCrLF + _
      "     					   EXISTS(SELECT str_id FROM tb_strutture WHERE (str_id = INSERTED.str_log_record OR (IsNull(INSERTED.str_log_record,0)=0  " + vbCrLF + _
      "                                                                          AND RegCode = INSERTED.str_log_CodAlb)) AND tipo = @AGRITURISMI_CON_RIC_TIP_ID))) " + vbCrLF + _
      "                          OR  " + vbCrLF + _
      "                          (Str_log_des LIKE '%cancellazione completa struttura%' AND (str_log_modello = @AGRITURISMI_MOD_ID OR  " + vbCrLF + _
      "                           EXISTS(SELECT str_log_id FROM tb_str_logs WHERE str_log_CodAlb = INSERTED.str_log_CodAlb AND IsNull(str_log_progressivo,0)>0))) " + vbCrLF + _
      "                          ) " + vbCrLF + _
      vbCrLF + _
      "     if (IsNull(@REGCODE,'') <> '') " + vbCRLF + _
      "         BEGIN " + vbCrLF + _
	  "				SELECT @REGCODE_REGIONE = regcode, @STR_ID_REGIONE = str_id FROM view_strutture WHERE RTRIM(LTRIM(IsNull(cod_ua_gestita,''))) LIKE @REGCODE AND IsNull(gestito_agenzia,0)=1 " + vbCrLf + _
	  "				IF (IsNull(@REGCODE_REGIONE, '')='') " + vbCrLf + _
	  "					BEGIN " + vbCrLf + _
	  "						SET @REGCODE_REGIONE = @REGCODE " + vbCrLF + _
	  "						SET @STR_ID_REGIONE = @STR_ID " + vbCrLF + _
	  "					END " + vbCrLf + _
	  vbCrLF + _
      "             IF (EXISTS(SELECT * FROM tb_str_logs WHERE RTRIM(LTRIM(str_log_CodAlb_regione)) LIKE @REGCODE_REGIONE AND str_log_id <> @LAST AND IsNull(str_log_progressivo,0)<>0 )) " + vbCrLf + _
      "                 BEGIN " + vbCrLF + _
      "                     SELECT @PROGRESSIVO = MAX(str_log_progressivo) FROM tb_str_logs " + vbCrLF + _
      "                         WHERE RTRIM(LTRIM(str_log_CodAlb_regione)) LIKE @REGCODE_REGIONE AND str_log_id <> @LAST " + vbCrLf + _
      "                     SET @PROGRESSIVO = @PROGRESSIVO + 1 " + vbCrLF + _
      "                 END " + vbCrLf + _
      "             ELSE " + vbCrLF + _
      "                 BEGIN " + vbCrLF + _
      "                     SET @PROGRESSIVO = 1 " + vbCrLF + _
      "                 END " + vbCrLF + _
      "             UPDATE tb_str_logs SET str_log_codalb_regione = @REGCODE_REGIONE, str_log_record_regione = @STR_ID_REGIONE, str_log_progressivo=@PROGRESSIVO WHERE str_log_id=@LAST " + vbCrLF + _
      "         END " + vbCrLF + _
      " ; "
CALL DB.Execute(sql, 1371)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1372
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__209(conn)
CALL DB.Execute(sql, 1372)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__209(conn)
end if
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1372)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1373
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__210(conn)
CALL DB.Execute(sql, 1373)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 1374
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__211(conn)
CALL DB.Execute(sql, 1374)
'*******************************************************************************************
'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1374)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1375
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__212(conn)
CALL DB.Execute(sql, 1375)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1376
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__213(conn)
CALL DB.Execute(sql, 1376)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1377
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__214(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1377, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1378
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__215(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1378, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1379
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__216(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1379, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1380
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__217(conn)
CALL DB.Execute(sql, 1380)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__217(conn)
end if
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1381
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__218(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1381, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1382
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__219(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1382, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1383
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__220(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1383, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1384
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__221(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1384, false, true)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1385
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__222(conn)
CALL DB.Execute(sql, 1385)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1386
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__223(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1386, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1387
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__224(conn)
CALL DB.Execute(sql, 1387)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__224(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1388
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__225(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1388, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1389
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__226(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1389, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1390
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__227(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1390, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1391
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__228(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1391, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1392
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__229(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1392, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1393
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__230(conn)
CALL DB.Execute(sql, 1393)
if DB.last_update_executed then
	CALL AggiornamentoSpeciale__FRAMEWORK_CORE__230(conn)
end if
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1394
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__231(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1394, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1395
'...........................................................................................
'	28/03/2013 - Nicola
'...........................................................................................
'	aggiunge campo per PEC
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD Lic_email_pec " & SQL_CharField(conn,60) & "  null; "
CALL DB.Execute(sql, 1395)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1396
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__232(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1396, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1397
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__233(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1397, false, false)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1398
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__234(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1398, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1399
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__235(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1399, false, true)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1400
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__207(conn)
CALL DB.Execute(sql, 1400)
'*******************************************************************************************

'*******************************************************************************************
CALL DB.SqlServer_VIEWS_REBUILD(1400)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1401
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__236(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1401, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1402
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__237(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1402, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1403
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__238(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1403, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1404
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__239(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1404, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1405
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__240(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1405, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1406
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__241(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1406, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1407
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD classifica_data SMALLDATETIME  null; "
CALL DB.ProtectedExecuteRebuild(sql, 1407, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1408
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__242(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1408, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1409
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__243(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1409, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1410
'...........................................................................................
sql = " ALTER TABLE tb_strutture ADD classifica_scadenza SMALLDATETIME  null; "
CALL DB.ProtectedExecuteRebuild(sql, 1410, false, true)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1411
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__244(conn)
CALL DB.ProtectedExecuteRebuild(sql, 1411, false, true)
'*******************************************************************************************


'COMMENTARE GLI AGGIONAMENTI DA QUA IN POI!!!




'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' INIZIO SERIE DI AGGIORNAMENTI PER FAR CONFLUIRE LE TRE SEZIONI DELLE UNITA' ABITATIVE,
' CON I RELATIVI DATI, IN UN UNICO APPLICATIVO  
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'AGGIORNAMENTO 1320
'...........................................................................................
'	22/12/2010 - Giacomo
'...........................................................................................
'	inserisce l'appicativo sul next-passport
'...........................................................................................
'sql = " INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_p1, sito_amministrazione, sito_prmEsterni_admin, sito_prmEsterni_sito) " & _
'	  " VALUES(" & TURISMO_UA & ",'Strutture ricettive [Unit&agrave;; abitative]','../Admin/UA','UA_USER',1,'../../Admin/Passport/PassportAdmin.asp','../../Admin/Passport/PassportSito.asp')"
'CALL DB.Execute(sql, 1320)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1321
'...........................................................................................
'	22/12/2010 - Giacomo
'...........................................................................................
'	inserisce i permessi agli amministratori per l'applicativo appena inserito, 
'	copiandoli dai permessi dei 3 applicativi delle UA
'...........................................................................................
'sql = " INSERT INTO rel_admin_sito(admin_id,sito_id,rel_as_permesso) " & _
'	  " SELECT admin_id, " & TURISMO_UA & ", 1 AS PERMESSO " & _
'	  " FROM rel_admin_sito WHERE sito_id IN (SELECT id_sito FROM tb_siti WHERE sito_nome LIKE '%unit%abitative%') "
'CALL DB.Execute(sql, 1321)
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO 1322
'...........................................................................................
'	22/12/2010 - Giacomo
'...........................................................................................
'	trasforma il modello 31 (U.A. NON CLASSIFICATE) nel nuovo modello UA
'...........................................................................................
'devo aggiornare anche Mod_Legge ????
'sql = " UPDATE tb_modelli SET Mod_Den_it = 'UNITA'' ABITATIVE AMMOBILIATE AD USO TURISTICO', Mod_strutture = 'U.A.', " & _
'	  " mod_Directory = 'UA', mod_applicazione_id = " & TURISMO_UA & ", Mod_Categoria_MAX = 3, " & _ 
'	  " mod_tipo_classificazione = '" & CLASSIFIED_BY_CATEGORY & "', " & _
'	  " mod_portale_pubblica = 1 " & _ 
'	  " WHERE Mod_Id = 31 "
'CALL DB.Execute(sql, 1322)
'*******************************************************************************************


'spostamento DOTAZIONI, SERVIZI e ZONE URBANE da U.A. gestite da agenzie a U.A. (EX NON CLASSIFICATE)
'*******************************************************************************************
'AGGIORNAMENTO 1323
'...........................................................................................
'	23/12/2010 - Giacomo
'...........................................................................................
'	spostamento DOTAZIONI del gruppo "CARATTERISTICHE UNITA' ABITATIVE"
'...........................................................................................
sql = " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 564 WHERE rel_str_id_dotaz = 671; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 563 WHERE rel_str_id_dotaz = 670; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 562 WHERE rel_str_id_dotaz = 669; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 559 WHERE rel_str_id_dotaz = 666; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 558 WHERE rel_str_id_dotaz = 665; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 560 WHERE rel_str_id_dotaz = 667; " & _
	  " UPDATE rel_str_dotaz SET rel_str_id_dotaz = 561 WHERE rel_str_id_dotaz = 668 "
'CALL DB.Execute(sql, 1323)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 1324
'...........................................................................................
'	23/12/2010 - Giacomo
'...........................................................................................
'	spostamento SERVIZI del gruppo "Impianti, attrezzature e servizi della struttura ricettiva"
'...........................................................................................
sql = " UPDATE rel_str_serv SET rel_str_id_relserv = 467 WHERE rel_str_id_relserv = 553; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 479 WHERE rel_str_id_relserv = 540; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 455 WHERE rel_str_id_relserv = 541; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 475 WHERE rel_str_id_relserv = 561; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 458 WHERE rel_str_id_relserv = 544; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 478 WHERE rel_str_id_relserv = 564; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 456 WHERE rel_str_id_relserv = 542; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 474 WHERE rel_str_id_relserv = 560; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 465 WHERE rel_str_id_relserv = 551; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 463 WHERE rel_str_id_relserv = 549; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 461 WHERE rel_str_id_relserv = 547; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 468 WHERE rel_str_id_relserv = 554; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 459 WHERE rel_str_id_relserv = 545; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 460 WHERE rel_str_id_relserv = 546; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 466 WHERE rel_str_id_relserv = 552; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 476 WHERE rel_str_id_relserv = 562; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 470 WHERE rel_str_id_relserv = 556; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 457 WHERE rel_str_id_relserv = 543; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 464 WHERE rel_str_id_relserv = 550; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 472 WHERE rel_str_id_relserv = 558; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 480 WHERE rel_str_id_relserv = 565; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 462 WHERE rel_str_id_relserv = 548; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 473 WHERE rel_str_id_relserv = 559; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 469 WHERE rel_str_id_relserv = 555; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 471 WHERE rel_str_id_relserv = 557; " & _
	  " UPDATE rel_str_serv SET rel_str_id_relserv = 477 WHERE rel_str_id_relserv = 563 "
'CALL DB.Execute(sql, 1324)
'*******************************************************************************************
'*******************************************************************************************
'AGGIORNAMENTO 1325
'...........................................................................................
'	23/12/2010 - Giacomo
'...........................................................................................
'	spostamento ZONE URBANE da U.A. gestite da agenzie a U.A. (EX NON CLASSIFICATE)
'...........................................................................................
sql = " UPDATE rel_zoneurb_str SET rel_zo_id = 116 WHERE rel_zo_id = 152; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 110 WHERE rel_zo_id = 142; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 105 WHERE rel_zo_id = 141; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 108 WHERE rel_zo_id = 146; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 113 WHERE rel_zo_id = 146; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 111 WHERE rel_zo_id = 151; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 106 WHERE rel_zo_id = 150; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 112 WHERE rel_zo_id = 149; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 109 WHERE rel_zo_id = 144; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 114 WHERE rel_zo_id = 145; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 107 WHERE rel_zo_id = 148; " & _
	  " UPDATE rel_zoneurb_str SET rel_zo_id = 115 WHERE rel_zo_id = 143 "
'CALL DB.Execute(sql, 1325)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1326
'...........................................................................................
'	23/12/2010 - Giacomo
'...........................................................................................
'	cambio riferimento dal modello da U.A. gestite da agenzie a U.A. (EX NON CLASSIFICATE) per le strutture
'...........................................................................................
sql = " UPDATE tb_loginStru SET Modello = 31 WHERE Modello = 36"
'CALL DB.Execute(sql, 1326)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1327
'...........................................................................................
'	23/12/2010 - Giacomo
'...........................................................................................
'	aggiungo permesso da Admin per l'utente NEXT-AIM sull'applicativo delle UA
'...........................................................................................
'sql = " INSERT INTO tb_turismo_admin_sito(tas_admin_id,tas_sito_id,tas_permesso,tas_ricezione_notifiche) " & _
'	  " VALUES(1," & TURISMO_UA & ",4,0) "
'CALL DB.Execute(sql, 1327)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1328
'...........................................................................................
'	'28/12/2010 - Giacomo
'...........................................................................................
'	correggo il nome di un tipo di struttura
'...........................................................................................
sql = "UPDATE tb_tipi_str SET Tip_Den_it = Replace(Tip_Den_it, 'à', 'a''')"
'CALL DB.Execute(sql, 1328)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1329
'...........................................................................................
'	'04/01/2011 - Giacomo
'...........................................................................................
'	cambio il riferimento dei log dalle U.A. GESTITE DA AGENZIE IMMOBILIARI (MOD_ID=36)
'	alle U.A. (MOD_ID=31)
'...........................................................................................
sql = "UPDATE tb_Str_logs SET str_log_modello = 31 WHERE str_log_modello = 36"
'CALL DB.Execute(sql, 1329)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 1330
'...........................................................................................
'	'04/01/2011 - Giacomo
'...........................................................................................
'	aggiungo la tipologia "Unità abitative classificate" per il modello 31 (U.A.)
'...........................................................................................
sql = " INSERT INTO tb_tipi_str(Tip_Mod_id,Tip_Den_it,Tip_valid_from,tip_cod_regione,tip_cod_RVT)" & _
	  " VALUES(31,'Unita'' abitative classificate',CONVERT(DATETIME, '1990-01-01 00:00:00', 102),0,0)"
'CALL DB.Execute(sql, 1330)
'*******************************************************************************************





'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' FINE SERIE DI AGGIORNAMENTI PER MODIFICA GESTIONE U.A.
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************************************************************************************************************************************************************************************************






'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************************************************************************************************************************
'AGGIORNAMENTO DI CHIUSURA ESEGUITO SEMPRE
'*********************************************************************************************************************************************************************************************************************************************************************************
DB.Terminate_SQL = _
        DropObject(conn, "VIEW_strutture", "VIEW") + _
        DropObject(conn, "VIEW_testata_strutture", "VIEW") + _
        DropObject(conn, "VIEW_valid_strutture", "VIEW") + _
        DropObject(conn, "VIEW_valid_testata_strutture", "VIEW") + _
        DropObject(conn, "VIEW_records_strutture", "VIEW") + _
        DropObject(conn, "VIEW_ALL_LOGIN", "VIEW") + _
        DropObject(conn, "VIEW_dichiarazioni_in_corso", "VIEW") + _
        DropObject(conn, "VIEW_Servizi", "VIEW") + _
	    _
        DropObject(conn, "fn_valid_strutture", "FUNCTION") + _
        DropObject(conn, "fn_records_strutture", "FUNCTION") + _
        _
        DropObject(conn, "spstr_UPDATE_tb_loginstru", "PROCEDURE") + _
        DropObject(conn, "spstr_DUPLICATE", "PROCEDURE") + _
        DropObject(conn, "spDuplicaStruttura", "PROCEDURE") + _
        DropObject(conn, "spNuovaStruttura", "PROCEDURE") + _
        DropObject(conn, "DELETE_Record_Struttura", "PROCEDURE") + _
        DropObject(conn, "spstr_WRITE_LOG", "PROCEDURE") + _
        DropObject(conn, "spstr_VALIDA_RECORD", "PROCEDURE") + _
        DropObject(conn, "spstr_DICHIARAZIONE_ANNULLA", "PROCEDURE") + _
        DropObject(conn, "spstr_DICHIARAZIONE_AVVISA", "PROCEDURE") + _
        DropObject(conn, "spstr_DICHIARAZIONE_COMPLETA", "PROCEDURE") + _
        DropObject(conn, "spstr_DICHIARAZIONE_PRESENTA", "PROCEDURE") + _
        DropObject(conn, "spstr_DICHIARAZIONE_RITIRA", "PROCEDURE") + _
        DropObject(conn, "spstr_CHECK_AND_DUPLICATE", "PROCEDURE") + _
        DropObject(conn, "spstr_CHECK_AND_DUPLICATE_4_ONLINE", "PROCEDURE")
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_Strutture AS " + vbCrLf + _
    	"   SELECT TOP 100 PERCENT " + vbCrLf + _
    	"       tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, dbo.tb_stru_gest.azienda_dati_altro, dbo.tb_stru_gest.azienda_dati_note, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND dbo.tb_loginStru.CURRENT_STR_ID = dbo.tb_strutture.Str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
        "   ORDER BY dbo.tb_strutture.Denominazione" + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_Testata_Strutture AS " + vbCrLf + _
        "   SELECT TOP 100 PERCENT  " + vbCrLf + _
        "       dbo.tb_strutture.str_ID, dbo.tb_strutture.Denominazione, tb_strutture.Lic_TipoSocieta, " + vbCrLf + _
        "       dbo.tb_strutture.RegCode, dbo.tb_strutture.Cod_Proprietario, dbo.tb_strutture.Cod_Tipologia, " + vbCrLF + _
        "       dbo.tb_strutture.Categoria, dbo.tb_strutture.Tipo, dbo.tb_strutture.Email, " + vbCrLf + _
        "       dbo.tb_strutture.Comune, dbo.tb_Strutture.PrezziEuro, tb_strutture.APTCode,  " + vbCrLf + _
        "       dbo.tb_strutture.DataModifica, dbo.tb_strutture.UtenteModifica, " + vbCrLf + _
        "       dbo.tb_strutture.online_modifica_data, dbo.tb_strutture.online_modifica_utente, " + vbCrLf + _
        "       dbo.tb_strutture.record_validato, dbo.tb_strutture.record_validato_data, dbo.tb_strutture.record_validato_utente, " + vbCrLf + _
        "       dbo.tb_strutture.avviso_inviato, dbo.tb_strutture.avviso_inviato_data, dbo.tb_strutture.avviso_inviato_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_presentata, dbo.tb_strutture.online_dic_presentata_data, dbo.tb_strutture.online_dic_presentata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_completata, dbo.tb_strutture.online_dic_completata_data, dbo.tb_strutture.online_dic_completata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_annullata, dbo.tb_strutture.online_dic_annullata_data, dbo.tb_strutture.online_dic_annullata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dichiarazione_id, dbo.tb_strutture.online_dic_tipo, dbo.tb_strutture.online_dic_anno_prezzi, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_data_inizio, dbo.tb_strutture.online_dic_data_fine, " + vbCrLf + _
        "       dbo.tb_strutture.anno_prezzi, archivio_modello_dichiarazione, archivio_tabella_prezzi, " + vbCrLf + _
        "       dbo.tb_strutture.nextInfo_area_id, " + vbCrLF + _
        "       dbo.tb_comuni.Comune AS ComuneTXT, tb_comuni.ufficio_apt, " + vbCrLf + _
        "       dbo.tb_tipi_str.Tip_Den_it, " + vbCrLf + _
        "       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.F_REVOCA_LIC, dbo.tb_stru_gest.F_REVOCA_CL,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_loginStru.*, " + vbCrLf + _
        "       dbo.tb_modelli.* " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND dbo.tb_loginStru.CURRENT_STR_ID = dbo.tb_strutture.Str_ID INNER JOIN  " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_id = dbo.tb_stru_gest.str_id INNER JOIN  " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_loginStru.modello = dbo.tb_modelli.Mod_id INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID " + vbCrLf + _
        "   ORDER BY dbo.tb_strutture.Denominazione" + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_valid_strutture AS " + vbCrLf + _
    	"   SELECT TOP 100 PERCENT " + vbCrLf + _
    	"       tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, dbo.tb_stru_gest.azienda_dati_altro, dbo.tb_stru_gest.azienda_dati_note, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND " + vbCrLf + _
		"		ISNULL(tb_loginStru.CURRENT_VALID_STR_ID, tb_loginStru.CURRENT_STR_ID) = dbo.tb_strutture.Str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
        "   ORDER BY dbo.tb_strutture.Denominazione" + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE FUNCTION dbo.fn_valid_strutture ( @anno int ) " + vbCrLF + _
		"   RETURNS table " + vbCrLF + _
		"   AS RETURN ( " + vbcRLF + _
		"   SELECT tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, dbo.tb_stru_gest.azienda_dati_altro, dbo.tb_stru_gest.azienda_dati_note, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND " + vbCrLf + _
		"                           str_id IN (SELECT TOP 1 str_id " + vbCrLf + _
		"                                      FROM tb_strutture tb_storico " + vbCrLF + _
		"                                      WHERE tb_storico.RegCode = tb_strutture.RegCode " + vbCrLF + _
		"                                            AND IsNull(tb_storico.anno_prezzi,Year(tb_storico.DataModifica))<=@anno " + vbCrLF + _
		"                                            AND ( str_id IN (SELECT str_id FROM tb_strutture tb_sub_storico " + vbCrLF + _
		"                                            				  WHERE tb_sub_storico.anno_prezzi = @anno AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                  OR " + vbCrLF + _
		"                                                  ( NOT EXISTS(SELECT str_id FROM tb_strutture tb_sub_storico " + vbCrLF + _
		"                                            					WHERE tb_sub_storico.anno_prezzi = @anno AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                    AND " + vbCrLF + _
		"                                                    str_id IN (SELECT str_id FROM tb_strutture tb_sub_storico " + vbCrLF + _
		"                                            					WHERE tb_sub_storico.DataModifica < CONVERT(DATETIME, CAST(@anno AS nvarchar(4)) + '-12-31 23:59:59', 102) " + vbCrLF + _
		"                                            						  AND tb_sub_storico.RegCode = tb_storico.Regcode) " + vbCrLF + _
		"                                                  ) " + vbCrLF + _
		"                                                ) " + vbCrlf + _
		"                                      ORDER BY tb_storico.regcode, tb_storico.DataModifica DESC, tb_storico.str_id DESC " + vbCrLF + _
		"                                     ) INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
		"   ) " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE FUNCTION dbo.fn_records_strutture ( @anno int ) " + vbCrLF + _
		"   RETURNS table " + vbCrLF + _
		"   AS RETURN ( " + vbcRLF + _
		"   SELECT tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, dbo.tb_stru_gest.azienda_dati_altro, dbo.tb_stru_gest.azienda_dati_note, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
		"	WHERE tb_strutture.str_id IN (SELECT TOP 1 str_id FROM tb_strutture " + vbCrLF + _
		"                                 WHERE RegCode = tb_strutture.RegCode AND Year(DataModifica)<=@ANNO ORDER BY DataModifica DESC, str_id DESC) " + vbCrLF + _
		"   ) " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_valid_testata_strutture AS " + vbCrLf + _
        "   SELECT TOP 100 PERCENT  " + vbCrLf + _
        "       dbo.tb_strutture.str_ID, dbo.tb_strutture.Denominazione, tb_strutture.Lic_TipoSocieta, " + vbCrLf + _
        "       dbo.tb_strutture.RegCode, dbo.tb_strutture.Cod_Proprietario, dbo.tb_strutture.Cod_Tipologia, " + vbCrLF + _
        "       dbo.tb_strutture.Categoria, dbo.tb_strutture.Tipo, dbo.tb_strutture.Email, " + vbCrLf + _
        "       dbo.tb_strutture.Comune, dbo.tb_Strutture.PrezziEuro, tb_strutture.APTCode,  " + vbCrLf + _
        "       dbo.tb_strutture.DataModifica, dbo.tb_strutture.UtenteModifica, " + vbCrLf + _
        "       dbo.tb_strutture.online_modifica_data, dbo.tb_strutture.online_modifica_utente, " + vbCrLf + _
        "       dbo.tb_strutture.record_validato, dbo.tb_strutture.record_validato_data, dbo.tb_strutture.record_validato_utente, " + vbCrLf + _
        "       dbo.tb_strutture.avviso_inviato, dbo.tb_strutture.avviso_inviato_data, dbo.tb_strutture.avviso_inviato_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_presentata, dbo.tb_strutture.online_dic_presentata_data, dbo.tb_strutture.online_dic_presentata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_completata, dbo.tb_strutture.online_dic_completata_data, dbo.tb_strutture.online_dic_completata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_annullata, dbo.tb_strutture.online_dic_annullata_data, dbo.tb_strutture.online_dic_annullata_utente, " + vbCrLf + _
        "       dbo.tb_strutture.online_dichiarazione_id, dbo.tb_strutture.online_dic_tipo, dbo.tb_strutture.online_dic_anno_prezzi, " + vbCrLf + _
        "       dbo.tb_strutture.online_dic_data_inizio, dbo.tb_strutture.online_dic_data_fine, " + vbCrLf + _
        "       dbo.tb_strutture.anno_prezzi, archivio_modello_dichiarazione, archivio_tabella_prezzi, " + vbCrLf + _
        "       dbo.tb_strutture.nextInfo_area_id, " + vbCrLF + _
        "       dbo.tb_comuni.Comune AS ComuneTXT, tb_comuni.ufficio_apt, " + vbCrLf + _
        "       dbo.tb_tipi_str.Tip_Den_it, " + vbCrLf + _
        "       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.F_REVOCA_LIC, dbo.tb_stru_gest.F_REVOCA_CL,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_loginStru.*, " + vbCrLf + _
        "       dbo.tb_modelli.* " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode AND " + vbCrLF + _
		"                           ISNULL(tb_loginStru.CURRENT_VALID_STR_ID, tb_loginStru.CURRENT_STR_ID) = dbo.tb_strutture.Str_ID INNER JOIN  " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_id = dbo.tb_stru_gest.str_id INNER JOIN  " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_loginStru.modello = dbo.tb_modelli.Mod_id INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID " + vbCrLf + _
        "   ORDER BY dbo.tb_strutture.Denominazione" + _
        " ; "

DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_records_strutture AS " + vbCrLf + _
    	"   SELECT TOP 100 PERCENT " + vbCrLf + _
    	"       tb_strutture.*,  " + vbCrLf + _
    	"       tb_modelli.*, " + vbCrLf + _
    	"       dbo.tb_loginStru.*, " + vbCrLf + _
		"		dbo.tb_tipi_str.*, " + VbCrLf + _
    	"       tb_comuni_lic.Comune AS Lic_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_closed.Comune AS closed_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_i_prop.Comune AS i_prop_COMUNE_TXT, " + vbCrLf + _
    	"       tb_comuni_i_loc.Comune AS i_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_prop.Comune AS a_prop_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_a_loc.Comune AS a_loc_COMUNETXT, " + vbCrLf + _
    	"       tb_comuni_RL.Comune AS rl_COMUNETXT, " + vbCrLf + _
    	"       dbo.tb_comuni.Comune AS COMUNETXT, " + vbCrLf + _
    	"       tb_comuni.ufficio_apt,  " + vbCrLf + _
    	"       dbo.tb_stru_gest.F_CH_TMP, dbo.tb_stru_gest.CH_TMP_IN, dbo.tb_stru_gest.CH_TMP_FI,  " + vbCrLf + _
        "       dbo.tb_stru_gest.CH_TMP_PROV, dbo.tb_stru_gest.CH_TMP_NUM, dbo.tb_stru_gest.F_REVOCA_LIC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_LIC, dbo.tb_stru_gest.REVOCA_LIC_PROV, dbo.tb_stru_gest.REVOCA_LIC_NUM,  " + vbCrLf + _
        "       dbo.tb_stru_gest.F_REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL, dbo.tb_stru_gest.REVOCA_CL_PROV,  " + vbCrLf + _
        "       dbo.tb_stru_gest.REVOCA_CL_NUM, dbo.tb_stru_gest.F_RIM_VINC, dbo.tb_stru_gest.RIM_VINC,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RIM_VINC_PROV, dbo.tb_stru_gest.RIM_VINC_NUM, dbo.tb_stru_gest.immobile_loc,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_nominativo, dbo.tb_stru_gest.i_prop_indirizzo, dbo.tb_stru_gest.i_prop_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_comune, dbo.tb_stru_gest.i_prop_cap, dbo.tb_stru_gest.i_prop_provincia, dbo.tb_stru_gest.i_prop_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_fax, dbo.tb_stru_gest.i_loc_nominativo, dbo.tb_stru_gest.i_loc_indirizzo, dbo.tb_stru_gest.i_loc_civico,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_comune, dbo.tb_stru_gest.i_loc_cap, dbo.tb_stru_gest.i_loc_provincia, dbo.tb_stru_gest.i_loc_telefono,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_loc_fax, dbo.tb_stru_gest.azienda_loc, dbo.tb_stru_gest.a_prop_nominativo, dbo.tb_stru_gest.a_prop_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_civico, dbo.tb_stru_gest.a_prop_comune, dbo.tb_stru_gest.a_prop_cap, dbo.tb_stru_gest.a_prop_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_telefono, dbo.tb_stru_gest.a_prop_fax, dbo.tb_stru_gest.a_loc_nominativo, dbo.tb_stru_gest.a_loc_indirizzo,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_civico, dbo.tb_stru_gest.a_loc_comune, dbo.tb_stru_gest.a_loc_cap, dbo.tb_stru_gest.a_loc_provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_loc_telefono, dbo.tb_stru_gest.a_loc_fax, dbo.tb_stru_gest.RL_cognome, dbo.tb_stru_gest.RL_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_indirizzo, dbo.tb_stru_gest.RL_civico, dbo.tb_stru_gest.RL_Comune, dbo.tb_stru_gest.RL_Provincia,  " + vbCrLf + _
        "       dbo.tb_stru_gest.RL_CAP, dbo.tb_stru_gest.RL_Telefono, dbo.tb_stru_gest.RL_Fax, dbo.tb_stru_gest.RL_Email,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_cognome, dbo.tb_stru_gest.i_prop_nome, dbo.tb_stru_gest.i_loc_cognome, dbo.tb_stru_gest.i_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.a_prop_cognome, dbo.tb_stru_gest.a_prop_nome, dbo.tb_stru_gest.a_loc_cognome, dbo.tb_stru_gest.a_loc_nome,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_data, dbo.tb_stru_gest.licenza_assegnata, dbo.tb_stru_gest.licenza_comune, dbo.tb_stru_gest.licenza_scadenza,  " + vbCrLf + _
        "       dbo.tb_stru_gest.licenza_rinnovo, dbo.tb_stru_gest.distintivo_Assegnato, dbo.tb_stru_gest.distintivo_Data,  " + vbCrLf + _
        "       dbo.tb_stru_gest.distintivo_restituzione, dbo.tb_stru_gest.abilitazione_data, dbo.tb_stru_gest.abilitazione_prov,  " + vbCrLf + _
        "       dbo.tb_stru_gest.abilitazione_ente, dbo.tb_stru_gest.a_prop_TipoSocieta, dbo.tb_stru_gest.a_loc_TipoSocieta,  " + vbCrLf + _
        "       dbo.tb_stru_gest.i_prop_TipoSocieta, dbo.tb_stru_gest.i_loc_TipoSocieta, dbo.tb_stru_gest.RL_CodFisc, dbo.tb_stru_gest.prov_tipo_1,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_1, dbo.tb_stru_gest.prov_data_1, dbo.tb_stru_gest.prov_ente_1, dbo.tb_stru_gest.prov_tipo_2,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_2, dbo.tb_stru_gest.prov_data_2, dbo.tb_stru_gest.prov_ente_2, dbo.tb_stru_gest.prov_tipo_3,  " + vbCrLf + _
        "       dbo.tb_stru_gest.prov_numero_3, dbo.tb_stru_gest.prov_data_3, dbo.tb_stru_gest.prov_ente_3, dbo.tb_stru_gest.azienda_dati_altro, dbo.tb_stru_gest.azienda_dati_note, " + vbCrLf + _
        "       dbo.tb_assoc.asc_nome " + vbCrLf + _
        "   FROM dbo.tb_loginStru INNER JOIN " + vbCrLF + _
        "       dbo.tb_strutture ON dbo.tb_loginStru.CODALB = dbo.tb_strutture.RegCode INNER JOIN " + vbCrLf + _
        "       dbo.tb_stru_gest ON dbo.tb_strutture.str_ID = dbo.tb_stru_gest.str_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_comuni ON dbo.tb_strutture.Comune = dbo.tb_comuni.Codice_ISTAT INNER JOIN " + vbCrLf + _
        "       dbo.tb_tipi_str ON dbo.tb_strutture.Tipo = dbo.tb_tipi_str.Tip_ID INNER JOIN " + vbCrLf + _
        "       dbo.tb_modelli ON dbo.tb_tipi_str.tip_Mod_ID = dbo.tb_modelli.Mod_ID LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_RL ON dbo.tb_stru_gest.RL_Comune = tb_comuni_RL.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_loc ON dbo.tb_stru_gest.a_loc_comune = tb_comuni_a_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_a_prop ON dbo.tb_stru_gest.a_prop_comune = tb_comuni_a_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_prop ON dbo.tb_stru_gest.i_prop_comune = tb_comuni_i_prop.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_closed ON dbo.tb_strutture.Closed_comune = tb_comuni_closed.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_lic ON dbo.tb_strutture.Lic_Comune = tb_comuni_lic.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_comuni tb_comuni_i_loc ON dbo.tb_stru_gest.i_loc_comune = tb_comuni_i_loc.Codice_ISTAT LEFT OUTER JOIN " + vbCrLf + _
        "       dbo.tb_assoc ON dbo.tb_strutture.associazione = tb_assoc.asc_id " + vbCrLf + _
        "   ORDER BY dbo.tb_strutture.RegCode" + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_ALL_LOGIN AS " + vbCrLf + _
        "   SELECT (CodAlb) AS RegCode, " + vbCrLF + _
        "       (0) AS asc_id, " + vbCrLF + _
        "       (UPPER(IsNull(Login, ''))) AS Login, " + vbCrLF + _
        "       (UPPER(IsNull([Password], ''))) AS [Password], " + vbCrLF + _
        "       (CASE WHEN mod_tipo_record='P' THEN 1 ELSE ISNULL(struttura_attiva, 0) END) AS login_attivo, " + vbCrLF + _
        "       modello " + vbCrLF + _
        "   FROM tb_loginStru INNER JOIN tb_modelli ON tb_loginstru.modello = tb_modelli.mod_id " + vbCrLF + _
        "   UNION " + vbCrLF + _
        "   SELECT ('') AS RegCode, " + vbCrLF + _
        "       (asc_id) AS asc_id, " + vbCrLF + _
        "       (UPPER(IsNull(asc_login, ''))) AS Login, " + vbCrLF + _
        "       (UPPER(IsNull(asc_password, ''))) AS [Password], " + vbCrLF + _
        "       (ISNULL(asc_Abilitato,0)) AS login_attivo, " + vbCrLF + _
        "       (0) AS modello " + vbCrLF + _
        "   FROM tb_assoc " + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_dichiarazioni_in_corso AS " + vbCrLf + _
	    "   SELECT * FROM tb_modelli INNER JOIN rel_dichiarazioni_modelli ON tb_modelli.mod_id = rel_dichiarazioni_modelli.rel_mod_id " + vbCrLf + _
	    "       INNER JOIN tb_dichiarazioni ON rel_dichiarazioni_modelli.rel_dic_id = tb_dichiarazioni.dic_id " + vbCrLf + _
	    "       WHERE CONVERT(DATETIME, CAST(Year(GetDate()) AS nvarchar(4)) + '-' + CAST(Month(GetDate()) AS nvarchar(2)) + '-' + CAST(Day(GetDate()) AS nvarchar(2)) + ' 00:00:00', 102) " + vbCrLF + _
		"			  BETWEEN tb_dichiarazioni.dic_data_inizio AND tb_dichiarazioni.dic_data_fine " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE VIEW dbo.VIEW_Servizi AS " + vbCrLf + _
	    "   SELECT tb_servizi.*, " + vbCrLf + _
	    "          rel_Grp_serv.rel_Grp_serv_id, tb_grp_vis.Grp_Mod_id, rel_Grp_serv.rel_Grp_serv_Ord, " + vbCrLf + _
	    "          (CASE WHEN GETDATE() BETWEEN rel_Grp_serv.serv_valid_from AND ISNULL(rel_Grp_serv.serv_valid_TO, GETDATE()+1) THEN 1 ELSE 0 END) AS VALIDO " + vbCrLf + _
	    "   FROM tb_servizi " + vbCrLf + _
	    "        INNER JOIN rel_Grp_serv ON tb_servizi.serv_id = rel_Grp_serv.rel_Grp_id_serv " + vbCrLf + _
	    "        INNER JOIN tb_grp_vis ON rel_Grp_serv.rel_id_Grp_serv = tb_grp_vis.Grp_id " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_UPDATE_tb_loginstru( " + vbCrLf + _
        "     @REGCODE nvarchar(12) " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "     DECLARE @current_str_id INT " + vbCrLf + _
        "     DECLARE @current_DataModifica SMALLDATETIME " + vbCrLF + _
        "     DECLARE @current_valid_str_id INT " + vbCrLf + _
        "     DECLARE @current_valid_DataModifica SMALLDATETIME " + vbCrLF + _
        "     DECLARE @current_dich_str_id INT " + vbCrLf + _
        "     DECLARE @current_dich_datamodifica SMALLDATETIME " + vbCrLf + _
        "     DECLARE @current_dich_id INT " + vbCrLf + _
        "     DECLARE @current_dich_tipo INT " + vbCrLf + _
        "     DECLARE @current_dich_anno_prezzi INT " + vbCrLf + _
        vbcrlf + _
        "     --recupera dati record corrente " + vbCrLf + _
        "     SELECT TOP 1 @current_str_id = str_id, " + vbCrLf + _
        "                  @current_DataModifica = DataModifica " + vbCrLf + _
        "         FROM tb_strutture WHERE RegCode=@REGCODE " + vbCrLf + _
        "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
        vbCrLf + _
        "     --recupera dati ultimo record validato " + vbCrLf + _
        "     SELECT TOP 1 @current_valid_str_id = str_id, " + vbCrLf + _
        "                  @current_valid_DataModifica = DataModifica " + vbCrLf + _
        "         FROM tb_strutture WHERE RegCode=@REGCODE AND IsNull(record_validato, 0)=1 " + vbCrLf + _
        "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
        vbCrLf + _
        "     --recupera dati ultimo record completato come dichiarazione " + vbCrLf + _
        "     SELECT TOP 1 @current_dich_str_id = str_id, " +  vbCrLf + _
        "                  @current_dich_datamodifica = DataModifica, " + vbCrLf + _
        "                  @current_dich_id = online_dichiarazione_id, " + vbCrLF + _
        "                  @current_dich_tipo = online_dic_tipo, " + vbCrLf + _
        "                  @current_dich_anno_prezzi = online_dic_anno_prezzi " + vbCrLf + _
        "         FROM tb_strutture WHERE RegCode=@REGCODE " + vbCrLf + _
        "                             AND IsNull(record_validato, 0)=1 " + vbCrLf + _
        "                             AND IsNull(online_dic_completata, 0)=1 " + vbCrLf+ _
        "                             AND IsNull(online_dic_annullata, 0)= 0 " + vbCrLf + _
        "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLf + _
        vbCrLf + _
        "     --aggiorna record tb_loginstru " + vbCrLf + _
        "     UPDATE tb_loginstru SET " + vbCrLf + _
        "         current_str_id = @current_str_id, " + vbCRlf + _
        "         current_DataModifica = @current_DataModifica, " + vbCrlf + _
        "         current_valid_str_id = @current_valid_str_id, " + vbCrLf + _
        "         current_valid_DataModifica = @current_valid_DataModifica, " + vbCrLf + _
        "         current_dich_str_id = @current_dich_str_id, " + vbCrLf + _
        "         current_dich_datamodifica = @current_dich_datamodifica, " + vbCrLf + _
        "         current_dich_id = @current_dich_id, " + vbCrLf + _
        "         current_dich_tipo = @current_dich_tipo, " + vbCrLf + _
        "         current_dich_anno_prezzi = @current_dich_anno_prezzi " + vbCrLf + _
        "     WHERE CodAlb = @RegCode " + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DUPLICATE( " + vbCrLF + _
        "     @STR_ID int, " + vbCrLF + _
        "     @DATA_NEW smalldatetime, " + vbCrLF + _
        "     @UTENTE nvarchar(50), " + vbCrLF + _
        "     @METODO nvarchar(10), --valori disponibili: VARIAZIONE_MODELLO, ONLINE, AMMINISTRAZIONE " + vbCrLF + _
        "     @NEW_STR_ID int OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "     DECLARE @RegCode nvarchar(12) " + vbCrLF + _
        "     --recupera codice regionale " + vbCrLF + _
        "     SELECT @RegCode = RegCode FROM tb_strutture WHERE str_id = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "     --duplica la struttura " + vbCrLF + _
        vbCrLf + _
        "     --inserimento nuovo record in tb_strutture " + vbCrLF + _
        "     IF (@METODO='VARIAZIONE_MODELLO') " + vbCrLF + _
        "     BEGIN " + vbCrLF + _
        "         --duplica la struttura per variazione modello: riporta tutti i dati tranne quelli della dichiarazione: " + vbCrLF + _
		"         --rimane lo storico con la situazione del modello contemporanea al periodo dichiarazione" + vbCrLF + _
        "         INSERT INTO tb_strutture( Denominazione, RegCode, DataModifica, UtenteModifica, Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, " + vbCrLF + _
		"                                   Settore, Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, " + vbCrLF + _
		"                                   Lic_DataNascita, Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, " + vbCrLF + _
		"                                   Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, " + vbCrLF + _
		"                                   Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, " + vbCrLF + _
		"                                   Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, " + vbCrLF + _
		"                                   prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, Inizio_Attivita, " + vbCrLF + _
		"                                   gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, note_interne, " + vbCrLF + _
		"                                   anno_prezzi, consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, " + vbCrLF + _
		"                                   record_validato, record_validato_data, record_validato_utente, avviso_inviato, avviso_inviato_data, avviso_inviato_utente, online_modifica_data, " + vbCrLF + _
		"                                   online_modifica_utente, online_dichiarazione_id, online_dic_tipo, online_dic_anno_prezzi, online_dic_data_inizio, online_dic_data_fine, " + vbCrLF + _
		"                                   online_dic_presentata, online_dic_presentata_data, online_dic_presentata_utente, online_dic_completata, online_dic_completata_data, " + vbCrLF + _
		"                                   online_dic_completata_utente, online_dic_annullata, online_dic_annullata_data, online_dic_annullata_utente, archivio_modello_dichiarazione, " + vbCrLF + _
		"                                   archivio_tabella_prezzi, nextInfo_area_id , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza )" + vbCrLF + _
        "             SELECT                Denominazione, RegCode, @DATA_NEW   , @UTENTE       , Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, " + vbCrLF + _
		"                                   Settore, Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, " + vbCrLF + _
		"                                   Lic_DataNascita, Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, " + vbCrLF + _
		"                                   Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, " + vbCrLF + _
		"                                   Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, " + vbCrLF + _
		"                                   Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, " + vbCrLF + _
		"                                   prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, Inizio_Attivita, " + vbCrLF + _
		"                                   gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, note_interne, " + vbCrLF + _
		"                                   anno_prezzi, consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, " + vbCrLF + _
		"                                   record_validato, record_validato_data, record_validato_utente, 0             , NULL               , NULL                 , NULL                , " + vbCrLF + _
		"                                   NULL                  , NULL                   , NULL           , NULL                  , NULL                  , NULL                , " + vbCrLF + _
		"                                   0                    , NULL                      , NULL                        , 0                    , NULL                      , " + vbCrLF + _
		"                                   NULL                        , 0                   , NULL                     , NULL                       , NULL                          , " + vbCrLF + _
		"                                   NULL                   , nextInfo_area_id , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza " + vbCrLF + _
        "             FROM tb_Strutture " + vbCrLF + _
        "             WHERE str_ID = @STR_ID " + vbCrLF + _
        "     END " + vbCrLF + _
        "     ELSE BEGIN " + vbCrLF + _
        "         --@METODO='ONLINE' OR @METODO='AMMINISTRAZIONE'" + vbCrLF + _
        "         DECLARE @DICHIARAZIONE_IN_CORSO_ID INT " + vbCrLF + _
        "         SELECT @DICHIARAZIONE_IN_CORSO_ID = (CASE WHEN custom_dichiarazione=1 THEN 0 ELSE dic_id END) " + vbCrLF + _
        "             FROM tb_loginstru LEFT JOIN VIEW_dichiarazioni_in_corso ON tb_loginstru.modello = VIEW_dichiarazioni_in_corso.mod_id  " + vbCrLF + _
        "             WHERE tb_loginstru.CodAlb = @RegCode " + vbCrLF + _
        vbCrLf + _
        "         DECLARE @DICHIARAZIONE_COMPILATA_ID INT " + vbCrLF + _
        "         DECLARE @DICHIARAZIONE_COMPLETATA BIT " + vbCrLF + _
        "         DECLARE @DICHIARAZIONE_ANNULLATA BIT " + vbCrLF + _
        "         SELECT @DICHIARAZIONE_COMPILATA_ID = online_dichiarazione_id, " + vbCrLF + _
        "                @DICHIARAZIONE_COMPLETATA = ISNULL(online_dic_completata,0), " + vbCrLF + _
        "                @DICHIARAZIONE_ANNULLATA = ISNULL(online_dic_annullata,0) " + vbCrLF + _
        "             FROM tb_strutture WHERE str_id = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "         IF ( @DICHIARAZIONE_COMPILATA_ID IS NULL OR " + vbCrLF + _
        "              @DICHIARAZIONE_IN_CORSO_ID <> @DICHIARAZIONE_COMPILATA_ID OR " + vbCrLF + _
        "              @DICHIARAZIONE_COMPLETATA=1 OR " + vbCrLF + _
        "              @DICHIARAZIONE_ANNULLATA=1) " + vbCrLF + _
        "         BEGIN " + vbCrLF + _
        "             --Duplica senza riportare dati dichiarazione: Fuori dichiarazione o Iter dichiarazione completata o dichiarazione annullata " + vbCrLF + _
        "             INSERT INTO tb_strutture( Denominazione, RegCode, DataModifica, UtenteModifica, Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, " + vbCrLF + _
		"                                   	Provincia, Settore, Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, " + vbCrLF + _
		"                                   	Lic_Nome, Lic_DataNascita, Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, " + vbCrLF + _
		"                                   	Lic_Telefono, Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, " + vbCrLF + _
		"                                   	Apertura, Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, " + vbCrLF + _
		"                                   	Closed_Ind, Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, " + vbCrLF + _
		"                                   	DecretoCL_N, DecretoCL_D, prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, " + vbCrLF + _
		"                                   	Inizio_Attivita, gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, " + vbCrLF + _
		"                                   	note_interne, anno_prezzi, consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, " + vbCrLF + _
		"                                   	lic_denominazione, record_validato, record_validato_data, record_validato_utente, avviso_inviato, avviso_inviato_data, avviso_inviato_utente, " + vbCrLF + _
		"                                   	online_modifica_data                                       , online_modifica_utente                                   , online_dichiarazione_id, " + vbCrLF + _
		"                                   	online_dic_tipo, online_dic_anno_prezzi, online_dic_data_inizio, online_dic_data_fine, online_dic_presentata, online_dic_presentata_data, " + vbCrLF + _
		"                                   	online_dic_presentata_utente, online_dic_completata, online_dic_completata_data, online_dic_completata_utente, online_dic_annullata, " + vbCrLF + _
		"                                   	online_dic_annullata_data, online_dic_annullata_utente, archivio_modello_dichiarazione, archivio_tabella_prezzi, nextInfo_area_id, " + vbCrLF + _
		"                                   	cod_ua_gestita, Registro_numero, Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza ) " + vbCrLF + _
        "                 SELECT                Denominazione, RegCode, @DATA_NEW   , @UTENTE       , Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, " + vbCrLF + _
		"                                   	Provincia, Settore, Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, " + vbCrLF + _
		"                                   	Lic_Nome, Lic_DataNascita, Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, " + vbCrLF + _
		"                                   	Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, Open_dal_1, " + vbCrLF + _
		"                                   	Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, Closed_civico, " + vbCrLF + _
		"                                   	Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, prezziEuro, " + vbCrLF + _
		"                                   	tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, Inizio_Attivita, gestito_agenzia, " + vbCrLF + _
		"                                   	interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, note_interne, anno_prezzi, " + vbCrLF + _
		"                                   	consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, 0              , " + vbCrLF + _
		"                                   	NULL                , NULL                  , 0             , NULL               , NULL                 , " + vbCrLF + _
		"                                   	CASE WHEN (@METODO = 'ONLINE') THEN @DATA_NEW ELSE NULL END, CASE WHEN (@METODO = 'ONLINE') THEN @UTENTE ELSE NULL END, NULL                   , " + vbCrLF + _
		"                                   	NULL           , NULL                  , NULL                  , NULL                , 0                    , NULL                      , " + vbCrLF + _
		"                                   	NULL                        , 0                    , NULL                      , NULL                        , 0                   , " + vbCrLF + _
		"                                   	NULL                     , NULL                       , NULL                          , NULL                   , nextInfo_area_id, " + vbCrLF + _
		"                                   	cod_ua_gestita, Registro_numero, Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza " + vbCrLF + _
        "                 FROM tb_Strutture " + vbCrLF + _
        "                 WHERE str_ID = @STR_ID " + vbCrLF + _
        "         END " + vbCrLF + _
        "         ELSE BEGIN " + vbCrLF + _
        "             --Duplica riportando dati dichiarazione: iter dichiarazione non ancora completato " + vbCrLF + _
        "             INSERT INTO tb_strutture( Denominazione, RegCode, DataModifica, UtenteModifica, Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, " + vbCrLF + _
		"                                   	Settore, Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, Lic_DataNascita, " + vbCrLF + _
		"                                   	Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, Lic_Fax, Lic_Email, Lic_CodFisc, " + vbCrLF + _
		"                                   	Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, " + vbCrLF + _
		"                                   	Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, " + vbCrLF + _
		"                                   	Closed_Fax, AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, " + vbCrLF + _
		"                                   	Piani_occupati, Piano_Immobile, Inizio_Attivita, gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, " + vbCrLF + _
		"                                   	foto_tesserino, recapito_presso, note_interne, anno_prezzi, consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, " + vbCrLF + _
		"                                   	closed_recapito_presso, lic_denominazione, record_validato, record_validato_data, record_validato_utente, avviso_inviato, avviso_inviato_data, " + vbCrLF + _
		"                                   	avviso_inviato_utente, online_modifica_data                                       , online_modifica_utente                                   , " + vbCrLF + _
		"                                   	online_dichiarazione_id, online_dic_tipo, online_dic_anno_prezzi, online_dic_data_inizio, online_dic_data_fine, online_dic_presentata, online_dic_presentata_data, " + vbCrLF + _
		"                                   	online_dic_presentata_utente, online_dic_completata, online_dic_completata_data, online_dic_completata_utente, online_dic_annullata, online_dic_annullata_data, " + vbCrLF + _
		"                                   	online_dic_annullata_utente, archivio_modello_dichiarazione, archivio_tabella_prezzi, nextInfo_area_id, cod_ua_gestita, Registro_numero, Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec ) " + vbCrLF + _
        "                 SELECT                Denominazione, RegCode, @DATA_NEW   , @UTENTE       , Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, Settore, " + vbCrLF + _
		"                                   	Telefono, Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, Lic_DataNascita, Lic_LuogoNascita, " + vbCrLF + _
		"                                   	Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, Anno_ristr, " + vbCrLF + _
		"                                   	Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, " + vbCrLF + _
		"                                   	Closed_Nome, Closed_Cognome, Closed_Ind, Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, " + vbCrLF + _
		"                                   	DecretoCL_N, DecretoCL_D, prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, Inizio_Attivita, " + vbCrLF + _
		"                                   	gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, note_interne, anno_prezzi, " + vbCrLF + _
		"                                   	consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, 0              , NULL                , " + vbCrLF + _
		"                                   	NULL                  , avviso_inviato, avviso_inviato_data, avviso_inviato_utente, CASE WHEN (@METODO = 'ONLINE') THEN @DATA_NEW ELSE NULL END, " + vbCrLF + _
		"                                   	CASE WHEN (@METODO = 'ONLINE') THEN @UTENTE ELSE NULL END, online_dichiarazione_id, online_dic_tipo, online_dic_anno_prezzi, online_dic_data_inizio, " + vbCrLF + _
		"                                   	online_dic_data_fine, online_dic_presentata, online_dic_presentata_data, online_dic_presentata_utente, online_dic_completata, online_dic_completata_data, " + vbCrLF + _
		"                                   	online_dic_completata_utente, 0                   , NULL                     , NULL                       , NULL                          , NULL                   , " + vbCrLF + _
		"                                   	nextInfo_area_id, cod_ua_gestita, Registro_numero, Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec " + vbCrLF + _
        "                 FROM tb_Strutture " + vbCrLF + _
        "                 WHERE str_ID = @STR_ID " + vbCrLF + _
        "         END " + vbCrLF + _
        "     END " + vbCrLF + _
        vbCrLf + _
        "     --legge id record inserito" + vbCrLF + _
        "     SELECT TOP 1 @NEW_STR_ID = str_id " + vbCrLF + _
        "         FROM tb_strutture " + vbCrLF + _
        "         WHERE RegCode=@REGCODE " + vbCrLF + _
        "         ORDER BY DataModifica DESC, str_id DESC " + vbCrLF + _
        vbCrLF + _
        "     --duplica tb_stru_gest " + vbCrLF + _
        "     INSERT INTO tb_stru_gest(str_ID,      RegCode, DataModifica, F_CH_TMP, CH_TMP_IN, CH_TMP_FI, CH_TMP_PROV, CH_TMP_NUM, F_REVOCA_LIC, REVOCA_LIC, REVOCA_LIC_PROV, " + vbCrLF + _
		"                              REVOCA_LIC_NUM, F_REVOCA_CL, REVOCA_CL, REVOCA_CL_PROV, REVOCA_CL_NUM, F_RIM_VINC, RIM_VINC, RIM_VINC_PROV, RIM_VINC_NUM, immobile_loc, " + vbCrLF + _
		"                              i_prop_nominativo, i_prop_indirizzo, i_prop_civico, i_prop_comune, i_prop_cap, i_prop_provincia, i_prop_telefono, i_prop_fax, i_loc_nominativo, " + vbCrLF + _
		"                              i_loc_indirizzo, i_loc_civico, i_loc_comune, i_loc_cap, i_loc_provincia, i_loc_telefono, i_loc_fax, azienda_loc, a_prop_nominativo, a_prop_indirizzo, " + vbCrLF + _
		"                              a_prop_civico, a_prop_comune, a_prop_cap, a_prop_provincia, a_prop_telefono, a_prop_fax, a_loc_nominativo, a_loc_indirizzo, a_loc_civico, " + vbCrLF + _
		"                              a_loc_comune, a_loc_cap, a_loc_provincia, a_loc_telefono, a_loc_fax, RL_cognome, RL_nome, RL_indirizzo, RL_civico, RL_Comune, RL_Provincia, " + vbCrLF + _
		"                              RL_CAP, RL_Telefono, RL_Fax, RL_Email, i_prop_cognome, i_prop_nome, i_loc_cognome, i_loc_nome, a_prop_cognome, a_prop_nome, a_loc_cognome, " + vbCrLF + _
		"                              a_loc_nome, licenza_data, licenza_assegnata, licenza_comune, licenza_scadenza, licenza_rinnovo, distintivo_Assegnato, distintivo_Data, " + vbCrLF + _
		"                              distintivo_restituzione, abilitazione_data, abilitazione_prov, abilitazione_ente, a_prop_TipoSocieta, a_loc_TipoSocieta, i_prop_TipoSocieta, " + vbCrLF + _
		"                              i_loc_TipoSocieta, RL_CodFisc, prov_tipo_1, prov_numero_1, prov_data_1, prov_ente_1, prov_tipo_2, prov_numero_2, prov_data_2, prov_ente_2, " + vbCrLF + _
		"                              prov_tipo_3, prov_numero_3, prov_data_3, prov_ente_3, azienda_dati_altro, azienda_dati_note) " + vbCrLF + _
        "         SELECT               @NEW_STR_ID, RegCode, @DATA_NEW   , F_CH_TMP, CH_TMP_IN, CH_TMP_FI, CH_TMP_PROV, CH_TMP_NUM, F_REVOCA_LIC, REVOCA_LIC, REVOCA_LIC_PROV, " + vbCrLF + _
		"                              REVOCA_LIC_NUM, F_REVOCA_CL, REVOCA_CL, REVOCA_CL_PROV, REVOCA_CL_NUM, F_RIM_VINC, RIM_VINC, RIM_VINC_PROV, RIM_VINC_NUM, immobile_loc, " + vbCrLF + _
		"                              i_prop_nominativo, i_prop_indirizzo, i_prop_civico, i_prop_comune, i_prop_cap, i_prop_provincia, i_prop_telefono, i_prop_fax, i_loc_nominativo, " + vbCrLF + _
		"                              i_loc_indirizzo, i_loc_civico, i_loc_comune, i_loc_cap, i_loc_provincia, i_loc_telefono, i_loc_fax, azienda_loc, a_prop_nominativo, a_prop_indirizzo, " + vbCrLF + _
		"                              a_prop_civico, a_prop_comune, a_prop_cap, a_prop_provincia, a_prop_telefono, a_prop_fax, a_loc_nominativo, a_loc_indirizzo, a_loc_civico, a_loc_comune, " + vbCrLF + _
		"                              a_loc_cap, a_loc_provincia, a_loc_telefono, a_loc_fax, RL_cognome, RL_nome, RL_indirizzo, RL_civico, RL_Comune, RL_Provincia, RL_CAP, RL_Telefono, " + vbCrLF + _
		"                              RL_Fax, RL_Email, i_prop_cognome, i_prop_nome, i_loc_cognome, i_loc_nome, a_prop_cognome, a_prop_nome, a_loc_cognome, a_loc_nome, licenza_data, " + vbCrLF + _
		"                              licenza_assegnata, licenza_comune, licenza_scadenza, licenza_rinnovo, distintivo_Assegnato, distintivo_Data, distintivo_restituzione, abilitazione_data, " + vbCrLF + _
		"                              abilitazione_prov, abilitazione_ente, a_prop_TipoSocieta, a_loc_TipoSocieta, i_prop_TipoSocieta, i_loc_TipoSocieta, RL_CodFisc, prov_tipo_1, prov_numero_1, " + vbCrLF + _
		"                              prov_data_1, prov_ente_1, prov_tipo_2, prov_numero_2, prov_data_2, prov_ente_2, prov_tipo_3, prov_numero_3, prov_data_3, prov_ente_3, azienda_dati_altro, " + vbCrLF + _
		"                              azienda_dati_note " + vbCrLF + _
        "         FROM tb_Stru_gest " + vbCrLF + _
        "         WHERE str_ID = @STR_ID " + vbCrLF + _
        vbCrLf + _
        "     --duplica unita' abitative " + vbCrLF + _
        "     INSERT INTO tb_ua(    id_struttura_ua, qta_ua, post_let_ua, nome_az_ua, muratura_ua, legno_ua, sintetico_ua, veranda_ua, angol_cott_ua, cucin_sep, num_vani, " + vbCrLF + _
		"                           pl_ssi_mobil_ua, pl_ssi_fisse_ua, pl_csi_mobil_ua, pl_csi_fisse_ua, pl_dis_mobil_ua, pl_dis_fisse_ua, prz_bs_min, prz_bs_max, prz_as_min, " + vbCrLF + _
		"                           prz_as_max, codice_ua, aria_cond_ua, riscaldamento_ua ) " + vbCrLF + _
        "         SELECT            @NEW_STR_ID    , qta_ua, post_let_ua, nome_az_ua, muratura_ua, legno_ua, sintetico_ua, veranda_ua, angol_cott_ua, cucin_sep, num_vani, " + vbCrLF + _
		"                           pl_ssi_mobil_ua, pl_ssi_fisse_ua, pl_csi_mobil_ua, pl_csi_fisse_ua, pl_dis_mobil_ua, pl_dis_fisse_ua, prz_bs_min, prz_bs_max, prz_as_min, " + vbCrLF + _
		"                           prz_as_max, codice_ua, aria_cond_ua, riscaldamento_ua " + vbCrLF + _
        "         FROM tb_ua " + vbCrLF + _
        "         WHERE id_struttura_ua = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "     --duplica replazioni con servizi " + vbCrLF + _
        "     INSERT INTO rel_str_serv(rel_str_id_relserv, rel_id_str_serv, rel_str_serv_val) " + vbCrLF + _
        "         SELECT               rel_str_id_relserv, @NEW_STR_ID    , rel_str_serv_val " + vbCrLF + _
        "         FROM rel_str_serv " + vbCrLF + _
        "         WHERE rel_id_str_serv = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "     --duplica relazioni con dotazioni " + vbCrLF + _
        "     INSERT INTO rel_str_dotaz(    rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_valore, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val) " + vbCrLF + _
        "         SELECT                    @NEW_STR_ID     , rel_str_id_dotaz, rel_str_dotaz_valore, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val " + vbCrLF + _
        "         FROM rel_str_dotaz " + vbCrLF + _
        "         WHERE rel_id_str_dotaz = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "     --duplica zone urbane " + vbCrLF + _
        "     INSERT INTO rel_zoneurb_str(rel_zo_id, rel_str_id , rel_val) " + vbCrLF + _
        "         SELECT                  rel_zo_id, @NEW_STR_ID, rel_val " + vbCrLF + _
        "         FROM rel_zoneurb_str " + vbCrLF + _
        "         WHERE rel_str_id = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "     --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato)" + vbCrLF + _
        "     EXEC spstr_UPDATE_tb_loginstru @RegCode " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spDuplicaStruttura( " + vbCrLF + _
        "     @CODREG nvarchar(12), " + vbCrLF + _
        "     @DATAREG smalldatetime, " + vbCrLF + _
        "     @NUOVADATA smalldatetime, " + vbCrLF + _
        "     @NEWSTRID int OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "     DECLARE @STRID int " + vbCrLF + _
        "     SELECT  @STRID = (SELECT MAX(str_ID) FROM dbo.tb_strutture " + vbCrLF + _
        "     WHERE (RegCode = @CODREG) AND (DataModifica = @DATAREG) GROUP BY RegCode) " + vbCrLF + _
        vbCrLF + _
        "     INSERT INTO dbo.tb_strutture(Denominazione, RegCode, DataModifica, Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, Settore, Telefono, " + vbCrLF + _
		"                                  Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, Lic_DataNascita, " + vbCrLF + _
		"                                  Lic_LuogoNascita, Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, Lic_Fax, Lic_Email, " + vbCrLF + _
		"                                  Lic_CodFisc, Anno_costr, Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, Open_dal_1, Open_al_1, Open_dal_2, " + vbCrLF + _
		"                                  Open_al_2, Open_dal_3, Open_al_3, Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, Closed_civico, Closed_comune, Closed_provincia, " + vbCrLF + _
		"                                  Closed_CAP, Closed_Telefono, Closed_Fax, AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, prezziEuro, tipoimmobile, note_compilazione, " + vbCrLF + _
		"                                  residenza_epoca, Occupazione_Immobile, Piani_occupati, Piano_Immobile, Inizio_Attivita, gestito_agenzia, interno, Cod_Proprietario, cellulare, " + vbCrLF + _
		"                                  Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, note_interne, anno_prezzi, consenso_foto, associazione, Terreno_MQ, " + vbCrLF + _
		"                                  Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, record_validato, record_validato_data, record_validato_utente, " + vbCrLF + _
		"                                  avviso_inviato, avviso_inviato_data, avviso_inviato_utente, online_modifica_data, online_modifica_utente, online_dichiarazione_id, online_dic_tipo, " + vbCrLF + _
		"                                  online_dic_anno_prezzi, online_dic_data_inizio, online_dic_data_fine, online_dic_presentata, online_dic_presentata_data, online_dic_presentata_utente, " + vbCrLF + _
		"                                  online_dic_completata, online_dic_completata_data, online_dic_completata_utente, online_dic_annullata, online_dic_annullata_data, " + vbCrLF + _
		"                                  online_dic_annullata_utente, archivio_modello_dichiarazione, archivio_tabella_prezzi, nextInfo_area_id, cod_ua_gestita, Registro_numero, " + vbCrLF + _
		"                                  Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza ) " + vbCrLF + _
        "         SELECT TOP 1             Denominazione, RegCode, @NUOVADATA  , Tipo, den_agg, Categoria, Indirizzo, civico, Localita, Comune, CAP, Altitudine, Provincia, Settore, Telefono, " + vbCrLF + _
		"                                  Fax, Telex, Sigla_telex, Email, WebUrl, resp_nome, resp_Cognome, Lic_Societa, Lic_TipoSocieta, Lic_Cognome, Lic_Nome, Lic_DataNascita, Lic_LuogoNascita, " + vbCrLF + _
		"                                  Lic_Numero, Lic_Data, Lic_Indirizzo, Lic_CIvico, Lic_Comune, Lic_Provincia, Lic_CAP, Lic_Telefono, Lic_Fax, Lic_Email, Lic_CodFisc, Anno_costr, " + vbCrLF + _
		"                                  Anno_ristr, Int_sto, Tipo_sto, Num_dip_fissi, Num_dip_stag, Num_mesi, Apertura, Open_dal_1, Open_al_1, Open_dal_2, Open_al_2, Open_dal_3, Open_al_3, " + vbCrLF + _
		"                                  Open_dal_4, Open_al_4, Closed_Nome, Closed_Cognome, Closed_Ind, Closed_civico, Closed_comune, Closed_provincia, Closed_CAP, Closed_Telefono, Closed_Fax, " + vbCrLF + _
		"                                  AutorizzSan_N, AutorizzSan_D, DecretoCL_N, DecretoCL_D, prezziEuro, tipoimmobile, note_compilazione, residenza_epoca, Occupazione_Immobile, Piani_occupati, " + vbCrLF + _
		"                                  Piano_Immobile, Inizio_Attivita, gestito_agenzia, interno, Cod_Proprietario, cellulare, Lic_Cellulare, consenso_pubblicazione, foto_tesserino, recapito_presso, " + vbCrLF + _
		"                                  note_interne, anno_prezzi, consenso_foto, associazione, Terreno_MQ, Cod_Tipologia, AptCode, closed_localita, closed_recapito_presso, lic_denominazione, " + vbCrLF + _
		"                                  record_validato, record_validato_data, record_validato_utente, avviso_inviato, avviso_inviato_data, avviso_inviato_utente, online_modifica_data, " + vbCrLF + _
		"                                  online_modifica_utente, online_dichiarazione_id, online_dic_tipo, online_dic_anno_prezzi, online_dic_data_inizio, online_dic_data_fine, online_dic_presentata, " + vbCrLF + _
		"                                  online_dic_presentata_data, online_dic_presentata_utente, online_dic_completata, online_dic_completata_data, online_dic_completata_utente, " + vbCrLF + _
		"                                  online_dic_annullata, online_dic_annullata_data, online_dic_annullata_utente, NULL                          , NULL                   , nextInfo_area_id, " + vbCrLF + _
		"                                  cod_ua_gestita, Registro_numero, Registro_data , Lic_PIva, Codice_Casa_Principale, Lic_email_pec, classifica_data, classifica_scadenza " + vbCrLF + _
        "         FROM tb_strutture " + vbCrLF + _
        "         WHERE (RegCode = @CODREG) AND (DataModifica = @DATAREG) " + vbCrLF + _
        "         ORDER BY DataModifica, Str_ID " + vbCrLF + _
        vbCrLF + _
        "     SELECT @NEWSTRID = @@IDENTITY " + vbCrLF + _
        vbCrLF + _
        "     INSERT INTO dbo.tb_stru_gest(str_ID   , RegCode, DataModifica, F_CH_TMP, CH_TMP_IN, CH_TMP_FI, CH_TMP_PROV, CH_TMP_NUM, F_REVOCA_LIC, REVOCA_LIC, REVOCA_LIC_PROV, REVOCA_LIC_NUM, " + vbCrLF + _
		"                                  F_REVOCA_CL, REVOCA_CL, REVOCA_CL_PROV, REVOCA_CL_NUM, F_RIM_VINC, RIM_VINC, RIM_VINC_PROV, RIM_VINC_NUM, immobile_loc, i_prop_nominativo, " + vbCrLF + _
		"                                  i_prop_indirizzo, i_prop_civico, i_prop_comune, i_prop_cap, i_prop_provincia, i_prop_telefono, i_prop_fax, i_loc_nominativo, i_loc_indirizzo, " + vbCrLF + _
		"                                  i_loc_civico, i_loc_comune, i_loc_cap, i_loc_provincia, i_loc_telefono, i_loc_fax, azienda_loc, a_prop_nominativo, a_prop_indirizzo, a_prop_civico, " + vbCrLF + _
		"                                  a_prop_comune, a_prop_cap, a_prop_provincia, a_prop_telefono, a_prop_fax, a_loc_nominativo, a_loc_indirizzo, a_loc_civico, a_loc_comune, a_loc_cap, " + vbCrLF + _
		"                                  a_loc_provincia, a_loc_telefono, a_loc_fax, RL_cognome, RL_nome, RL_indirizzo, RL_civico, RL_Comune, RL_Provincia, RL_CAP, RL_Telefono, RL_Fax, " + vbCrLF + _
		"                                  RL_Email, i_prop_cognome, i_prop_nome, i_loc_cognome, i_loc_nome, a_prop_cognome, a_prop_nome, a_loc_cognome, a_loc_nome, licenza_data, licenza_assegnata, " + vbCrLF + _
		"                                  licenza_comune, licenza_scadenza, licenza_rinnovo, distintivo_Assegnato, distintivo_Data, distintivo_restituzione, abilitazione_data, abilitazione_prov, " + vbCrLF + _
		"                                  abilitazione_ente, a_prop_TipoSocieta, a_loc_TipoSocieta, i_prop_TipoSocieta, i_loc_TipoSocieta, RL_CodFisc, prov_tipo_1, prov_numero_1, prov_data_1, " + vbCrLF + _
		"                                  prov_ente_1, prov_tipo_2, prov_numero_2, prov_data_2, prov_ente_2, prov_tipo_3, prov_numero_3, prov_data_3, prov_ente_3, azienda_dati_altro, azienda_dati_note) " + vbCrLF + _
        "         SELECT TOP 1             @NEWSTRID, RegCode, @NUOVADATA  , F_CH_TMP, CH_TMP_IN, CH_TMP_FI, CH_TMP_PROV, CH_TMP_NUM, F_REVOCA_LIC, REVOCA_LIC, REVOCA_LIC_PROV, REVOCA_LIC_NUM, " + vbCrLF + _
		"                                  F_REVOCA_CL, REVOCA_CL, REVOCA_CL_PROV, REVOCA_CL_NUM, F_RIM_VINC, RIM_VINC, RIM_VINC_PROV, RIM_VINC_NUM, immobile_loc, i_prop_nominativo, i_prop_indirizzo, " + vbCrLF + _
		"                                  i_prop_civico, i_prop_comune, i_prop_cap, i_prop_provincia, i_prop_telefono, i_prop_fax, i_loc_nominativo, i_loc_indirizzo, i_loc_civico, i_loc_comune, " + vbCrLF + _
		"                                  i_loc_cap, i_loc_provincia, i_loc_telefono, i_loc_fax, azienda_loc, a_prop_nominativo, a_prop_indirizzo, a_prop_civico, a_prop_comune, a_prop_cap, " + vbCrLF + _
		"                                  a_prop_provincia, a_prop_telefono, a_prop_fax, a_loc_nominativo, a_loc_indirizzo, a_loc_civico, a_loc_comune, a_loc_cap, a_loc_provincia, a_loc_telefono, " + vbCrLF + _
		"                                  a_loc_fax, RL_cognome, RL_nome, RL_indirizzo, RL_civico, RL_Comune, RL_Provincia, RL_CAP, RL_Telefono, RL_Fax, RL_Email, i_prop_cognome, i_prop_nome, " + vbCrLF + _
		"                                  i_loc_cognome, i_loc_nome, a_prop_cognome, a_prop_nome, a_loc_cognome, a_loc_nome, licenza_data, licenza_assegnata, licenza_comune, licenza_scadenza, " + vbCrLF + _
		"                                  licenza_rinnovo, distintivo_Assegnato, distintivo_Data, distintivo_restituzione, abilitazione_data, abilitazione_prov, abilitazione_ente, a_prop_TipoSocieta, " + vbCrLF + _
		"                                  a_loc_TipoSocieta, i_prop_TipoSocieta, i_loc_TipoSocieta, RL_CodFisc, prov_tipo_1, prov_numero_1, prov_data_1, prov_ente_1, prov_tipo_2, prov_numero_2, " + vbCrLF + _
		"                                  prov_data_2, prov_ente_2, prov_tipo_3, prov_numero_3, prov_data_3, prov_ente_3, azienda_dati_altro, azienda_dati_note " + vbCrLF + _
        "         FROM tb_stru_gest " + vbCrLF + _
        "         WHERE (str_ID = @STRID) " + vbCrLF + _
        vbCrLF + _
        "     INSERT INTO dbo.rel_str_dotaz (rel_id_str_dotaz, rel_str_id_dotaz, rel_str_dotaz_valore, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val) " + vbCrLF + _
        "         SELECT                     @NEWSTRID       , rel_str_id_dotaz, rel_str_dotaz_valore, rel_str_dotaz_testo_it, rel_str_dotaz_pos_val " + vbCrLF + _
        "         FROM dbo.rel_str_dotaz " + vbCrLF + _
        "         WHERE (rel_id_str_dotaz = @STRID) " + vbCrLF + _
        vbCrLF + _
        "     INSERT INTO dbo.rel_str_serv (rel_str_id_relserv, rel_id_str_serv, rel_str_serv_val) " + vbCrLF + _
        "         SELECT                    rel_str_id_relserv, @NEWSTRID      , rel_str_serv_val " + vbCrLF + _
        "         FROM dbo.rel_str_serv " + vbCrLF + _
        "         WHERE (rel_id_str_serv = @STRID) " + vbCrLF + _
        vbCrLF + _
        "     INSERT INTO dbo.rel_zoneurb_str (rel_zo_id, rel_str_id, rel_val) " + vbCrLF + _
        "         SELECT                       rel_zo_id, @NEWSTRID , rel_val " + vbCrLF + _
        "         FROM dbo.rel_zoneurb_str " + vbCrLF + _
        "         WHERE (rel_str_id = @STRID) " + vbCrLF + _
        vbCrLF + _
        "     --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato)" + vbCrLF + _
        "     EXEC spstr_UPDATE_tb_loginstru @CODREG " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE  PROCEDURE dbo.spNuovaStruttura ( " + vbCrLF + _
        "   @DENOMINAZIONE nvarchar(60), " + vbCrLF + _
        "   @CODALB nvarchar(12), " + vbCrLF + _
        "   @TIPOLOGIA int, " + vbCrLF + _
        "   @DATA_MOD smalldatetime, " + vbCrLF + _
        "   @CODCOM nvarchar(6), " + vbCrLF + _
        "   @MODELLO int, " + vbCrLF + _
        "   @STRID int OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "   INSERT INTO dbo.tb_loginstru (CodAlb , Login, Password, modello , struttura_attiva, custom_dichiarazione) " + vbCrLF + _
        "       VALUES                   (@CODALB, ''   , ''      , @MODELLO, 0               , 0                   ) " + vbCrLF + _
        vbCrLF + _
        "   INSERT INTO dbo.tb_strutture (Denominazione , RegCode, Tipo      , DataModifica, prezziEuro, Comune , record_validato, avviso_inviato, online_dic_presentata, " + vbCrLF + _
		"                                  online_dic_completata, online_dic_annullata, online_dichiarazione_id ) " + vbCrLF + _
        "       VALUES                   (@DENOMINAZIONE, @CODALB, @TIPOLOGIA, @DATA_MOD   , 1         , @CODCOM, 0              , 0             , 0                    , " + vbCrLF + _
		"                                  0                    , 0                   , NULL                    ) " + vbCrLF + _
        vbCrLF + _
        "   SELECT @STRID = @@IDENTITY " + vbCrLF + _
        "   INSERT INTO dbo.tb_stru_gest (str_ID, RegCode, DataModifica) " + vbCrLF + _
        "       VALUES                   (@STRID, @CODALB, @DATA_MOD   ) " + vbCrLF + _
        vbCrLf + _
        "   --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato) " + vbCrLF + _
        "   EXEC spstr_UPDATE_tb_loginstru @CODALB " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE Procedure dbo.DELETE_Record_Struttura ( " + vbCrLF + _
        "   @STR_ID int " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "   DECLARE @REGCODE nvarchar(12) " + vbCrLF + _
        "   SELECT @REGCODE = RegCode FROM tb_Strutture WHERE Str_ID=@STR_ID " + vbCrLF + _
        vbCrLF + _
        "   DELETE FROM rel_zoneurb_str WHERE rel_str_id = @STR_ID " + vbCrLF + _
        "   DELETE FROM rel_str_serv WHERE rel_id_str_serv = @STR_ID " + vbCrLF + _
        "   DELETE FROM rel_str_dotaz WHERE rel_id_str_dotaz = @STR_ID " + vbCrLF + _
        "   DELETE FROM tb_ua WHERE id_struttura_ua = @STR_ID " + vbCrLF + _
        "   DELETE FROM tb_stru_gest WHERE str_id = @STR_ID " + vbCrLF + _
        "   DELETE FROM tb_strutture WHERE str_ID = @STR_ID " + vbCrLF + _
        vbCrLF + _
        "   --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato) " + vbCrLF + _
        "   EXEC spstr_UPDATE_tb_loginstru @REGCODE " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_WRITE_LOG( " + vbCrLF + _
        "   @STR_ID int, " + vbCrLF + _
        "   @REGCODE nvarchar(12), " + vbCrLF + _
        "   @LOGIN nvarchar(50), " + vbCrLF + _
        "   @OPERAZIONE nvarchar(50), " + vbCrLF + _
        "   @UPDATE_LOG_RECORD bit " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLF + _
        "   IF ((@REGCODE='') OR (@REGCODE IS NULL)) " + vbCrLF + _
        "   BEGIN " + vbCrLF + _
        "       SELECT @REGCODE=RegCode FROM tb_strutture WHERE Str_ID=@STR_ID " + vbCrLF + _
        "   END " + vbCrLF + _
        vbCrLF + _
        "   IF (@STR_ID=0) " + vbCrLF + _
        "   BEGIN " + vbCrLF + _
        "       SELECT TOP 1 @STR_ID = Str_ID FROM tb_strutture WHERE RegCode = @REGCODE ORDER BY DataModifica DESC, str_id DESC" + vbCrLF + _
        "   END " + vbCrLF + _
        vbCrLF + _
		"   DECLARE @MODELLO int " + vbCrLf + _
		"   SELECT TOP 1 @MODELLO = modello FROM tb_loginstru WHERE CodAlb LIKE LEFT(@REGCODE, 3) + '%' " + vbCrLF + _
		vbCrLf + _
        "   INSERT INTO tb_str_logs (Str_log_CodAlb, Str_log_data, Str_log_ope  , Str_log_des       , Str_log_record, str_log_modello) " + vbCrLF + _
        "       VALUES              (@REGCODE      , GETDATE()   , RTRIM(@LOGIN), RTRIM(@OPERAZIONE), @STR_ID,        @MODELLO) " + vbCrLF + _
        vbCrLF + _
        "   IF (@UPDATE_LOG_RECORD=1) " + vbCrLF + _
        "   BEGIN " + vbCrLF + _
        "       UPDATE tb_strutture SET UtenteModifica=RTRIM(@LOGIN) WHERE str_id=@STR_ID " + vbCrLF + _
        "   END " + vbCrLF + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE Procedure dbo.spstr_VALIDA_RECORD ( " + vbCrLf + _
        "   @STR_ID int, " + vbCrLf + _
        "   @UTENTE nvarchar(50), " + vbCrLf + _
        "   @ANNO_PREZZI int " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   DECLARE @REGCODE nvarchar(12) " + vbCrLf + _
        "   SELECT @REGCODE = RegCode FROM tb_Strutture WHERE Str_ID=@STR_ID " + vbCrLf + _
        vbCrLf + _
        "   --imposta dati su record struttura " + vbCrLf + _
        "   UPDATE tb_strutture SET record_validato=1, " + vbCrLf + _
        "                           record_validato_data = GETDATE(), " + vbCrLf + _
        "                           record_validato_utente = @Utente, " + vbCrLf + _
        "                           anno_prezzi = @ANNO_PREZZI " + vbCrLf + _
        "       WHERE str_id=@STR_ID " + vbCrLf + _
        vbCrLf + _
        "   --aggiorna tb_loginStru e porta puntamento ai nuovi record (record corrente e record validato)" + vbCrLf + _
        "   EXEC spstr_UPDATE_tb_loginstru @REGCODE " + vbCrLf + _
        vbCrLf + _
        "   --scrive record su log " + vbCrLf + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Validita dati: registrazione validata.', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DICHIARAZIONE_RITIRA( " + vbCrLf + _
        "   @RegCode nvarchar(12), " + vbCrLf + _
        "   @STR_ID INT, " + vbCrLf + _
        "   @UTENTE nvarchar(50) " + vbCrLf + _
        " ) " + vbCrLF + _
        " AS " + vbCrLf + _
        "   --imposta dati su record struttura: annulla dichiarazione e rimuove validazione" + vbCrLf + _
        "   UPDATE tb_strutture SET " + vbCrLF + _
        "           online_dic_annullata = 1, " + vbCrLf + _
        "           online_dic_annullata_data = GETDATE(), " + vbCrLf + _
        "           online_dic_annullata_utente = @UTENTE, " + vbCrLf + _
        "           record_validato=0, " + vbCrLf + _
        "           record_validato_data = NULL, " + vbCrLf + _
        "           record_validato_utente = NULL, " + vbCrLf + _
        "           anno_prezzi = NULL " + vbCrLf + _
        "       WHERE str_id = @STR_ID " + vbCRLF + _
        vbCrLf + _
        "   --aggiorna dati tb_loginstru " + vbCRLF + _
        "   EXEC spstr_UPDATE_tb_loginstru @REGCODE " + vbCrLf + _
        vbCrLf + _
        "   --registra riga su log " + vbCrLF + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Stato dichiarazione: dichiarazione ritirata', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DICHIARAZIONE_ANNULLA( " + vbCrLf + _
        "   @RegCode nvarchar(12), " + vbCrLf + _
        "   @STR_ID INT, " + vbCrLf + _
        "   @UTENTE nvarchar(50), " + vbCrLf + _
        "   @STR_ID_NEW INT OUTPUT " + vbCrLF + _
        " ) " + vbCrLF + _
        " AS " + vbCrLf + _
        "   --imposta dati su record struttura: annulla dichiarazione e rimuove validazione" + vbCrLf + _
        "   UPDATE tb_strutture SET " + vbCrLF + _
        "           online_dic_annullata = 1, " + vbCrLf + _
        "           online_dic_annullata_data = GETDATE(), " + vbCrLf + _
        "           online_dic_annullata_utente = @UTENTE, " + vbCrLf + _
        "           record_validato=0, " + vbCrLf + _
        "           record_validato_data = NULL, " + vbCrLf + _
        "           record_validato_utente = NULL, " + vbCrLf + _
        "           anno_prezzi = NULL, " + vbCrLf + _
		"			archivio_modello_dichiarazione = NULL, " + vbCrLF + _
		"           archivio_tabella_prezzi = NULL " + vbCrLf + _
        "       WHERE str_id = @STR_ID " + vbCRLF + _
        vbCrLf + _
        "   --aggiorna dati tb_loginstru " + vbCRLF + _
        "   EXEC spstr_UPDATE_tb_loginstru @REGCODE " + vbCrLf + _
        vbCrLf + _
        "   --registra riga su log " + vbCrLF + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Stato dichiarazione: dichiarazione annullata', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DICHIARAZIONE_AVVISA( " + vbCrLf + _
        "   @RegCode nvarchar(12), " + vbCrLf + _
        "   @STR_ID INT, " + vbCrLf + _
        "   @UTENTE nvarchar(50) " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   UPDATE tb_strutture SET avviso_inviato=1, " + vbCrLf + _
        "                           avviso_inviato_data = GETDATE(), " + vbCrLf + _
        "                           avviso_inviato_utente = @UTENTE " + vbCrLf + _
        "   WHERE str_id = @STR_ID " + vbCrLf + _
        vbCrLf + _
        "   --scrive record su log " + vbCrLf + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Stato dichiarazione: avviso di completamento inviato', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DICHIARAZIONE_COMPLETA( " + vbCrLf + _
        "   @RegCode nvarchar(12), " + vbCrLf + _
        "   @STR_ID INT, " + vbCrLf + _
        "   @DICH_ID INT, " + vbCrLf + _
        "   @DICH_INIZIO smalldatetime, " + vbCrLf + _
        "   @DICH_FINE smalldatetime, " + vbCrLf + _
        "   @DICH_TIPO int, " + vbCrLf + _
        "   @DICH_ANNO_PREZZI int, " + vbCrLf + _
        "   @UTENTE nvarchar(50) " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   --esegue presentazione della dichiarazione se non ancora eseguita " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_PRESENTATA bit " + vbCrLf + _
        "   SELECT @DICHIARAZIONE_PRESENTATA = IsNull(online_dic_presentata,0) FROM tb_strutture WHERE str_id = @STR_ID " + vbCrLf + _
        "   IF (@DICHIARAZIONE_PRESENTATA=0) " + vbCrLf + _
        "       EXEC spstr_DICHIARAZIONE_PRESENTA @REGCODE, @STR_ID, @DICH_ID, @DICH_INIZIO, @DICH_FINE, @DICH_TIPO, @DICH_ANNO_PREZZI, @UTENTE " + vbCrLf + _
        vbCrLf + _
        "   UPDATE tb_strutture SET online_dic_completata=1, " + vbCrLf + _
        "                           online_dic_completata_data = GETDATE(), " + vbCrLf + _
        "                           online_dic_completata_utente = @UTENTE, " + vbCrLf + _
        "                           online_dichiarazione_id = @DICH_ID, " + vbCrLf + _
        "                           online_dic_tipo = @DICH_TIPO, " + vbCrLf + _
        "                           online_dic_anno_prezzi = @DICH_ANNO_PREZZI, " + vbCrLf + _
        "                           online_dic_data_inizio = @DICH_INIZIO, " + vbCrLf + _
        "                           online_dic_data_fine = @DICH_FINE, " + vbCrLf + _
        "                           online_dic_annullata = 0, " + vbCrLf + _
        "                           online_dic_annullata_data = NULL, " + vbCrLf + _
        "                           online_dic_annullata_utente = NULL " + vbCrLf + _
        "       WHERE str_id = @STR_ID " + vbCrLf + _
        vbCrLf + _
        "   EXEC spstr_VALIDA_RECORD @STR_ID, @UTENTE, @DICH_ANNO_PREZZI " + vbCrLf + _
        vbCRLf + _
        "   --scrive record su log " + vbCrLf + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Stato dichiarazione: dichiarazione completata e validata', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_DICHIARAZIONE_PRESENTA( " + vbCrLf + _
        "   @RegCode nvarchar(12), " + vbCrLf + _
        "   @STR_ID INT, " + vbCrLf + _
        "   @DICH_ID INT, " + vbCrLf + _
        "   @DICH_INIZIO smalldatetime, " + vbCrLf + _
        "   @DICH_FINE smalldatetime, " + vbCrLf + _
        "   @DICH_TIPO int, " + vbCrLf + _
        "   @DICH_ANNO_PREZZI int, " + vbCrLf + _
        "   @UTENTE nvarchar(50) " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   UPDATE tb_strutture SET online_dic_presentata=1, " + vbCrLf + _
        "                           online_dic_presentata_data = GETDATE(), " + vbCrLf + _
        "                           online_dic_presentata_utente = @UTENTE, " + vbCrLf + _
        "                           online_dichiarazione_id = @DICH_ID, " + vbCrLf + _
        "                           online_dic_tipo = @DICH_TIPO, " + vbCrLf + _
        "                           online_dic_anno_prezzi = @DICH_ANNO_PREZZI, " + vbCrLf + _
        "                           online_dic_data_inizio = @DICH_INIZIO, " + vbCrLf + _
        "                           online_dic_data_fine = @DICH_FINE, " + vbCrLf + _
        "                           online_dic_completata = 0, " + vbCrLf + _
        "                           online_dic_completata_data = NULL, " + vbCrLf + _
        "                           online_dic_completata_utente = NULL, " + vbCrLf + _
        "                           online_dic_annullata = 0, " + vbCrLf + _
        "                           online_dic_annullata_data = NULL, " + vbCrLf + _
        "                           online_dic_annullata_utente = NULL " + vbCrLf + _
        "   WHERE str_id = @STR_ID " + vbCrLf + _
        vbCrLf + _
        "   --scrive record su log " + vbCrLf + _
        "   EXEC spstr_WRITE_LOG @STR_ID, @REGCODE, @UTENTE, 'Stato dichiarazione: dichiarazione presentata', 0 " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_CHECK_AND_DUPLICATE( " + vbCrLf + _
        "   @STR_ID int, " + vbCrLf + _
        "   @UTENTE nvarchar(50), " + vbCrLf + _
        "   @NEW_STR_ID int OUTPUT " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   DECLARE @REGCODE nvarchar(12) " + vbCrLf + _
        "   DECLARE @DATA_NEW smalldatetime " + vbCrLf + _
        "   DECLARE @DATA_OLD smalldatetime " + vbCrLf + _
        "   DECLARE @RECORD_VALIDATO bit " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_COMPLETATA bit " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_ANNULLATA BIT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_COMPILATA_ID INT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_ID INT " + vbCrLf + _
        vbCRLf + _
        "   --recupera dati registrazione corrente " + vbCrLf + _
        "   SELECT @REGCODE = RegCode, " + vbCrLf + _
        "          @DATA_OLD = DataModifica, " + vbCrLf + _
        "          @RECORD_VALIDATO = record_validato, " + vbCrLf + _
        "          @DICHIARAZIONE_COMPILATA_ID = online_dichiarazione_id, " + vbCrLf + _
        "          @DICHIARAZIONE_COMPLETATA = online_dic_completata, " + vbCrLf + _
        "          @DICHIARAZIONE_ANNULLATA = online_dic_annullata " + vbCrLf + _
        "       FROM tb_Strutture WHERE str_ID=@STR_ID " + vbCrLf + _
        vbCrLf + _
        "   --calcola data corrente " + vbCrLf + _
        "   SET @DATA_NEW = GETDATE() " + vbCrLf + _
        "   set @DATA_NEW = CONVERT(DATETIME, str(YEAR(@DATA_NEW)) + '-' + str(MONTH(@DATA_NEW)) + '-' + str(DAY(@DATA_NEW)) + ' 00:00:00', 102) " + vbCrLf + _
        vbCrLf + _
        "   SELECT @DICHIARAZIONE_IN_CORSO_ID = (CASE WHEN custom_dichiarazione=1 THEN 0 ELSE dic_id END) " + vbCrLF + _
        "       FROM tb_loginstru LEFT JOIN VIEW_dichiarazioni_in_corso ON tb_loginstru.modello = VIEW_dichiarazioni_in_corso.mod_id  " + vbCrLF + _
        "       WHERE tb_loginstru.CodAlb = @RegCode " + vbCrLF + _
        vbCrLf + _
        "   --verifica criteri di duplicazione: data modifica diversa da oggi, record validato o di dichiarazione confermata " + vbCrLf + _
        "   IF (    (   (YEAR(@DATA_NEW) = YEAR(@DATA_OLD)) AND " + vbCrLf + _
        "               (MONTH(@DATA_NEW) = MONTH(@DATA_OLD)) AND " + vbCrLf + _
        "               (DAY(@DATA_NEW) = DAY(@DATA_OLD))   ) " + vbCrLf + _
        "       AND @RECORD_VALIDATO=0 " + vbCrLf + _
        "       AND @DICHIARAZIONE_COMPLETATA=0 " + vbCrLf + _
        "       AND @DICHIARAZIONE_ANNULLATA=0 " + vbCrLf + _
        "       AND (   @DICHIARAZIONE_COMPILATA_ID IS NULL OR " + vbCrLf + _
        "               @DICHIARAZIONE_COMPILATA_ID = @DICHIARAZIONE_IN_CORSO_ID    )   ) " + vbCrLf + _
        "   BEGIN " + vbCrLf + _
        "       --data uguale e record non validato ne confermato o annullato da dichiarazione online " + vbCrLf + _
        "       SET @NEW_STR_ID = @STR_ID " + vbCrLf + _
        "   END " + vbCrLf + _
        "   ELSE BEGIN " + vbCrLf + _
        "       --record da duplicare perche' cambiata la data di modifica, o record valiato o di dichiarazione online confermata " + vbCrLf + _
        "       EXEC spstr_DUPLICATE @STR_ID, @DATA_NEW, @UTENTE, 'AMMINISTRAZIONE', @NEW_STR_ID OUTPUT " + vbCrLf + _
        "   END " + vbCrLf + _
        " ; "
DB.Terminate_SQL = DB.Terminate_SQL + _
        " CREATE PROCEDURE dbo.spstr_CHECK_AND_DUPLICATE_4_ONLINE( " + vbCrLf + _
        "   @REGCODE nvarchar(12), " + vbCrLf + _
        "   @UTENTE nvarchar(50), " + vbCrLf + _
        "   @NEW_STR_ID int OUTPUT " + vbCrLf + _
        " ) " + vbCrLf + _
        " AS " + vbCrLf + _
        "   DECLARE @STR_ID int " + vbCrLf + _
        "   DECLARE @DATA_ONLINE_NEW smalldatetime " + vbCrLf + _
        "   DECLARE @DATA_ONLINE_OLD smalldatetime " + vbCrLf + _
        "   DECLARE @RECORD_VALIDATO bit " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_COMPLETATA bit " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_ANNULLATA BIT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_COMPILATA_ID INT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_ID INT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_TIPO INT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_ANNO_PREZZI INT " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_INIZIO smalldatetime " + vbCrLf + _
        "   DECLARE @DICHIARAZIONE_IN_CORSO_FINE smalldatetime " + vbCrLf + _
        vbCrLF + _
        "   --recupera calcola nuova data " + vbCrLf + _
        "   SET @DATA_ONLINE_NEW = GETDATE() " + vbCrLF + _
        "   set @DATA_ONLINE_NEW = CONVERT(DATETIME, str(YEAR(@DATA_ONLINE_NEW)) + '-' + str(MONTH(@DATA_ONLINE_NEW)) + '-' + str(DAY(@DATA_ONLINE_NEW)) + ' 00:00:00', 102) " + vbcRlf + _
        vbCrLf + _
        "   --recupera dati record corrente " + vbCrLf + _
        "   SELECT TOP 1 @STR_ID=str_ID, " + vbCrLf + _
        "                @DATA_ONLINE_OLD = online_modifica_data, " + vbCrLf + _
        "                @RECORD_VALIDATO = record_validato, " + vbCrLf + _
        "                @DICHIARAZIONE_COMPILATA_ID = online_dichiarazione_id, " + vbCrLf + _
        "                @DICHIARAZIONE_COMPLETATA = online_dic_completata, " + vbCrLf + _
        "                @DICHIARAZIONE_ANNULLATA = online_dic_annullata " + vbCrLf + _
        "       FROM View_testata_strutture WHERE RegCode=@REGCODE " +  vbCrLf + _
        vbCrLF + _
        "   --recupera termini dichiarazione in corso " + vbCrLf + _
        "   SELECT @DICHIARAZIONE_IN_CORSO_ID = (CASE WHEN custom_dichiarazione=1 THEN 0 ELSE dic_id END), " + vbCrLF + _
        "          @DICHIARAZIONE_IN_CORSO_INIZIO = (CASE WHEN custom_dichiarazione=1 THEN custom_dic_data_inizio ELSE dic_data_inizio END), " + vbCrLF + _
        "          @DICHIARAZIONE_IN_CORSO_FINE = (CASE WHEN custom_dichiarazione=1 THEN custom_dic_data_fine ELSE dic_data_fine END) " + vbCrLF + _
        "       FROM tb_loginstru LEFT JOIN VIEW_dichiarazioni_in_corso ON tb_loginstru.modello = VIEW_dichiarazioni_in_corso.mod_id  " + vbCrLF + _
        "       WHERE tb_loginstru.CodAlb = @RegCode " + vbCrLF + _
        vbCrLF + _
        "   --verifiche pro-duplicazione " + vbcRlf + _
        "   IF ( (@DATA_ONLINE_OLD IS NULL) OR " + vbCrLf + _
        "        (@RECORD_VALIDATO = 1) OR " + vbCrLF + _
        "        (@DICHIARAZIONE_COMPLETATA = 1) OR " + vbCrLf + _
        "        (@DICHIARAZIONE_ANNULLATA = 1) OR " + vbCrLf + _
        "        (@DICHIARAZIONE_COMPILATA_ID IS NULL) OR " + vbCrLf + _
        "        (@DICHIARAZIONE_COMPILATA_ID <> @DICHIARAZIONE_IN_CORSO_ID)  ) " + vbCrLf + _
        "   BEGIN " + vbCrLF + _
        "       --DUPLICA RECORD: " + vbCrLF + _
        "       --modifica on-line non ancora effettuata o compilazione non impostata o record validato o " + vbCrLF + _
        "       --compilazione completata o compilazione annullata o dichiarazione in corso diversa da quella del record (se presente) " + vbCRLf + _
        "       SET @NEW_STR_ID = 0 " + vbCrLf + _
        "   END " + vbCrLf + _
        "   ELSE BEGIN " + vbCrLF + _
        "       IF ( (@DATA_ONLINE_NEW = @DATA_ONLINE_OLD) OR " + vbCrLF + _
        "            ( (@DATA_ONLINE_OLD BETWEEN @DICHIARAZIONE_IN_CORSO_INIZIO AND @DICHIARAZIONE_IN_CORSO_FINE) AND " + vbCrLf + _
        "              (@DATA_ONLINE_NEW BETWEEN @DICHIARAZIONE_IN_CORSO_INIZIO AND @DICHIARAZIONE_IN_CORSO_FINE) )  ) " + vbCrLF + _
        "       BEGIN " + vbCrLf + _
        "           --NON DUPLICA RECORD: " + vbCRLf + _
        "           --data uguale o date di modifica entrambe dentro il periodo di dichiarazione " + vbCrLf + _
        "           SET @NEW_STR_ID = @STR_ID " + vbCrLf + _
        "       END " + vbCrLf + _
        "       ELSE BEGIN " + vbCrLF + _
        "           --DUPLICA RECORD " + vbCRLf + _
        "           --data online diversa o una delle date non e' nel periodo di dichiarazione corrente " + vbCRLf + _
        "           SET @NEW_STR_ID = 0 " + vbCrLf + _
        "       END " + vbCrLf + _
        "   END " + vbCRLf + _
        vbCrLf + _
        "   IF ( @NEW_STR_ID=0 ) " + vbCrLf + _
        "   BEGIN " + vbCrLf + _
        "       -- duplica la struttura " + vbCrLf + _
        "       EXEC spstr_DUPLICATE @STR_ID, @DATA_ONLINE_NEW, @UTENTE, 'ONLINE', @NEW_STR_ID OUTPUT " + vbCrLf + _
        "   END " + vbCrLF + _
        vbCrLF + _
        "   --aggiorna dati registrazione ed imposta dati compilazione" + vbCrLf + _
        "   UPDATE tb_strutture SET DataModifica = @DATA_ONLINE_NEW, " + vbCrLf + _
        "                           UtenteModifica = @UTENTE, " + vbCRLf + _
        "                           online_modifica_data = @DATA_ONLINE_NEW, " + vbCrLf + _
        "                           online_modifica_utente = @UTENTE, " + vbCrLF + _
        "                           online_dichiarazione_id = @DICHIARAZIONE_IN_CORSO_ID, " + vbCrLf + _
        "                           online_dic_tipo = @DICHIARAZIONE_IN_CORSO_TIPO, " + vbCrLf + _
        "                           online_dic_anno_prezzi = @DICHIARAZIONE_IN_CORSO_ANNO_PREZZI, " + vbCRLF + _
        "                           online_dic_data_inizio = @DICHIARAZIONE_IN_CORSO_INIZIO, " + vbCrLF + _
        "                           online_dic_data_fine = @DICHIARAZIONE_IN_CORSO_FINE " + vbCrLF + _
        "       WHERE str_id = @NEW_STR_ID " + vbcRLF + _
        vbCrLF + _
        "   UPDATE tb_stru_gest SET DataModifica = @DATA_ONLINE_NEW " + vbCrLF + _
        "       WHERE str_id = @NEW_STR_ID " + vbcRLF + _
        " ; "
'*******************************************************************************************



%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
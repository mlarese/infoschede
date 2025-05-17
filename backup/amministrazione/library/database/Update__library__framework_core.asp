
<!--#INCLUDE FILE="../../NextWeb5/Tools_NextWeb5_extension.asp" -->
<%

'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti comuni del NEXT-FRAMEWORK
'per tutte le istanze
'APPLICATIVI COMPRESI:
'	NEXT-PASSPORT
'	NEXT-COM / NEXT-DOC
'	NEXT-NEWS
'	NEXT-LINK
'	NEXT-FAQ
'	NEXT-GALLERY
'	NEXT-TEAM
' 	NEXT-WEB 5.0
'...........................................................................................
'...........................................................................................


'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'FUNZIONI PER LA MANUTENZIONE DEL FRAMEWORK CORE

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************


'*******************************************************************************************
'Aggiornamento e ripristino dei nomi corretti delle applicazioni
'...........................................................................................
function rebuild__FRAMEWORK_CORE__Nomi_Applicazioni(conn)
	rebuild__FRAMEWORK_CORE__Nomi_Applicazioni = _
		"UPDATE tb_siti SET sito_nome='NEXT-passport [gestione utenti]', sito_dir='NextPassport' WHERE id_sito=" & NEXTPASSPORT & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-web [gestione grafica e contenuti]', sito_dir='NextWeb' WHERE id_sito=" & NEXTWEB & ";" + _
		"UPDATE tb_siti SET sito_nome='" + IIF(Application("NextCrm"), "NEXT-doc+ [comunicazioni & documenti]", "NEXT-com [gestione comunicazioni]") + "', sito_dir='NextCom' WHERE id_sito=" & NEXTCOM & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-news [gestione news]', sito_dir='NextNews' WHERE id_sito=" & NEXTNEWS & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-link [gestione link utili]', sito_dir='NextLink' WHERE id_sito=" & NEXTLINK & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-menu [gestione men&ugrave;; e ricette]', sito_dir='NextMenu' WHERE id_sito=" & NEXTMENU & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-flat [gestione prenotazioni appartamenti turistici]', sito_dir='NextFlat' WHERE id_sito=" & NEXTFLAT & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-memo [gestione pubblicazione documenti]', sito_dir='NextMemo' WHERE id_sito=" & NEXTMEMO & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-banner [gestione banners pubblicitari]', sito_dir='NextBanner' WHERE id_sito=" & NEXTBANNER & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-club [gestione associati]', sito_dir='NextClub' WHERE id_sito=" & NEXTCLUB & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-booking [gestione prenotazioni]', sito_dir='NextBooking' WHERE id_sito=" & NEXTBOOKING & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-guestbook [gestione guestbook]', sito_dir='NextGuestbook' WHERE id_sito=" & NEXTGUESTBOOK & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-contract [gestione bandi ed appalti]', sito_dir='NextContract' WHERE id_sito=" & NEXTCONTRACT & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-f.a.q. [gestione frequently asked questions]', sito_dir='NextFaq' WHERE id_sito=" & NEXTFAQ & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-team [gestione organigramma aziendale]', sito_dir='NextTeam' WHERE id_sito=" & NEXTTEAM & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-booking portal [gestione portale di prenotazione]', sito_dir='NextBookingPortal' WHERE id_sito=" & NEXTBOOKINGPORTALE & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-flat portal [gestione portale di prenotazione appartamenti]', sito_dir='NextFlatPortal' WHERE id_sito=" & NEXTFLATPORTAL & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-realestate [gestione immobili]', sito_dir='NextRealEstate' WHERE id_sito=" & NEXTREALESTATE & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-b2b [gestione prodotti, magazzino e vendita]', sito_dir='NextB2b' WHERE id_sito=" & NEXTB2B & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-school [gestione strutture scolastiche]', sito_dir='NextSchool' WHERE id_sito=" & NEXTSCHOOL & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-congress [gestione spazi congressuali]', sito_dir='NextCongress' WHERE id_sito=" & NEXTCONGRESS & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-b2b integration [Import, export ed integrazione dati]', sito_dir='../NextB2B_Integration' WHERE id_sito=" & NEXTB2B_IMPORT & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-b2b mailing [Statistiche ed elaborazioni per mailing list]', sito_dir='NextB2b_Mailing' WHERE id_sito=" & NEXTB2B_MAILING & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-travel [gestione agenzia di viaggi]', sito_dir='NextTravel' WHERE id_sito=" & NEXTTRAVEL & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-web 4.0 [gestione grafica e contenuti]', sito_dir='NextWeb4' WHERE id_sito=" & NEXTWEB4 & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-web 5.0 [gestione grafica e contenuti accessibili]', sito_dir='NextWeb5' WHERE id_sito=" & NEXTWEB5 & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-gallery [gestione gallerie di immagini]', sito_dir='NextGallery' WHERE id_sito=" & NEXTGALLERY & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-booking 2.0 [gestione prenotazioni]', sito_dir='NextBooking2' WHERE id_sito=" & NEXTBOOKING2 & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-info [gestione eventi ed informazioni turistiche]', sito_dir='NextInfo' WHERE id_sito=" & NEXTINFO & ";" + _
		"UPDATE tb_siti SET sito_nome='NEXT-booking 3.0 [gestione prenotazioni]', sito_dir='NextBooking3' WHERE id_sito=" & NEXTBOOKING3 & ";" + _
        "UPDATE tb_siti SET sito_nome='NEXT-tour [gestione pacchetti turistici]', sito_dir='NextTour' WHERE id_sito=" & NEXTTOUR & ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'Aggiornamento speciale 
'...........................................................................................
sub rebuild__FRAMEWORK_CORE__cartelle(conn, rs, DB, version)
	dim sql, fso, FolderUpload, FolderSite, FolderTemp, FolderSiteDocs
	'esegue un aggiornamento fasullo per aumentare la versione
	sql = "SELECT * FROM AA_Versione"
	CALL DB.Execute(sql, version)
	if DB.last_update_executed then
		'esegue aggiornamento solo se la versione e' corretta e la query fasulla e' stata eseguita
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set FolderUpload = fso.GetFolder(Application("IMAGE_PATH"))
	
		'ripulisce directory temporanee
		CALL ClearTempDir(fso)
		
		'rimuove cartella template (presente solo in alcuni)
		CALL FolderRemove(fso, Application("IMAGE_PATH") & "\template", false)
	
		'rimuove file inutili dalle directory (qualsiasi directory)
		CALL FileRemove(fso, Application("IMAGE_PATH"), "thumbs.db", true)
		CALL FileRemove(fso, Application("IMAGE_PATH"), "pspbrwse.jbf", true)
		
		'verifica esistenza delle cartella docs principale
		if not fso.FolderExists(FolderUpload.path + "\docs") then
			CALL fso.CreateFolder(FolderUpload.path + "\docs")
		end if
		
		'verifica esistenza delle cartella temp principale
		if not fso.FolderExists(FolderUpload.path + "\temp") then
			CALL fso.CreateFolder(FolderUpload.path + "\temp")
		end if
		
		'recupera directory temporanea per generazione diretory utenti
		set FolderTemp = fso.GetFolder(FolderUpload.path + "\temp")
		if not fso.FolderExists(FolderTemp.path + "\docs") then
			CALL fso.CreateFolder(FolderTemp.path + "\docs")
		end if
		
		set FolderTemp = fso.GetFolder(FolderUpload.path + "\temp\docs")
		sql = "SELECT * FROM tb_admin"
		rs.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
		while not rs.eof
			if not fso.FolderExists(FolderTemp.path + "\" & rs("admin_login")) then
				CALL fso.CreateFolder(FolderTemp.path + "\" & rs("admin_login"))
			end if
			rs.movenext
		wend
		rs.close
		
		'corregge directory del next web
		sql = "SELECT id_sito FROM tb_siti WHERE id_sito IN (" & NEXTWEB & ", " & NEXTWEB4 & ", " & NEXTWEB5 & ")"
		rs.open sql, conn, adOpenstatic, adLockOptimistic, adCmdText
		if not rs.eof then
			'scorre tutte le directory dei siti (solo con nome numerico)
			for each FolderSite in FolderUpload.SubFolders
				if isNumeric(FolderSite.name) then
					'rimuove cartella exports
					CALL FolderRemove(fso, FolderSite.path & "\exports", false)
					
					'rimuove cartella temp
					CALL FolderRemove(fso, FolderSite.path & "\temp", false)
					
					'rimuove cartelle vuote da docs interna (eventualmente anche docs)
					CALL RemoveEmptyFolders(fso, FolderSite.path + "\docs")
					
					'controllo: se esiste ancora vuol dire che contiene files / cartelle che devono essere spostati.
					if fso.FolderExists(FolderSite.path + "\docs") then
						set FolderSiteDocs = fso.GetFolder(FolderSite.path + "\docs")
						'rinomina eventuali cartele emails da "<id>" a "eml_<id>"
						for each SubFolder in FolderSiteDocs.SubFolders
							if IsNumeric(SubFolder.name) then
								'directory non vuota che contiene i files delle email: deve essere rinominata
								SubFolder.name = "eml_" + SubFolder.name
							elseif instr(1, SubFolder, "pra_", vbTextCompare)>0 then
								'cancella directory residua delle pratiche
								CALL SubFolder.Delete(true)
							end if
						next
						'Copia Directory upload/<az_id>/docs su upload/docs
						CALL FolderSiteDocs.Copy(FolderUpload.path + "\docs", true)
						'rimuove vecchia directory Docs
						CALL FolderSiteDocs.Delete(true)
					end if
					
					if rs("id_sito") = NEXTWEB4 OR rs("id_sito") = NEXTWEB5 then
						'rimuove cartella oggetti
						CALL FolderRemove(fso, FolderSite.path & "\objects", false)
				
						'rimuove file vuoto.jpg e/o obj_vuoto.jpg nelle varie cartelle
						CALL FileRemove(fso, FolderSite.path & "\flash", "vuoto.jpg", true)
						CALL FileRemove(fso, FolderSite.path & "\flash", "obj_vuoto.jpg", true)
						
						CALL FileRemove(fso, FolderSite.path & "\images", "vuoto.jpg", true)
						CALL FileRemove(fso, FolderSite.path & "\images", "obj_vuoto.jpg", true)
					end if
					
				end if
			next
		end if
		rs.close
	end if
end sub




'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'FUNZIONI PER L'INSTALLAZIONE DELLE COMPONENTI CORE DEL FRAMEWORK PER DATABASE SQL

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'*******************************************************************************************
'INSTALLA CORE DEL FRAMEWORK TEMPLATE
'...........................................................................................
function Install__FRAMEWORK_CORE(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE = ""
		case DB_SQL
			Install__FRAMEWORK_CORE = _
				Install__FRAMEWORK_CORE__NEXTPASSPORT(conn) + _
	  			Install__FRAMEWORK_CORE__NEXTCOM(conn) + _
	  			Install__FRAMEWORK_CORE__NEXTLINK(conn) + _
				Install__FRAMEWORK_CORE__NEXTNEWS(conn) + _
				Install__FRAMEWORK_CORE__NEXTGALLERY(conn) + _
				Install__FRAMEWORK_CORE__NEXTFAQ(conn) + _
				Install__FRAMEWORK_CORE__NEXTTEAM(conn) + _
				Install__FRAMEWORK_CORE__NEXTWEB5(conn)
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-PASSPORT
'...........................................................................................
'aggiunge il NEXT-PASSPORT. anche se la funzione e' separata e' necessario installare anche
'il NEXTcom perche' tb_utenti e' collegata a tb_indirizzario tramite relazione creata in
'Install__FRAMEWORK_CORE__NEXTCOM.
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTPASSPORT(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTPASSPORT = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTPASSPORT = _
				" CREATE TABLE dbo.tb_siti ( " + _
				"		id_sito int NOT NULL , " + _
				"		sito_nome nvarchar (250) NULL , " + _
				"		sito_dir nvarchar (150) NULL , " + _
				"		sito_p1 nvarchar (50) NULL , " + _
				"		sito_p2 nvarchar (50) NULL , " + _
				"		sito_p3 nvarchar (50) NULL , " + _
				"		sito_p4 nvarchar (50) NULL , " + _
				"		sito_p5 nvarchar (50) NULL , " + _
				"	 	sito_p6 nvarchar (50) NULL , " + _
				"		sito_p7 nvarchar (50) NULL , " + _
				"		sito_p8 nvarchar (50) NULL , " + _
				"		sito_p9 nvarchar (50) NULL , " + _
				"		sito_amministrazione bit NULL , " + _
				"		sito_rubrica_area_riservata int NULL " + _
				" ); " + _
				" ALTER TABLE tb_siti WITH NOCHECK ADD CONSTRAINT PK_tb_siti PRIMARY KEY NONCLUSTERED (id_sito); " + _
				" CREATE TABLE dbo.tb_siti_parametri ( " + _
				"		par_id int IDENTITY (1, 1) NOT NULL , " + _
				"		par_key nvarchar (250) NULL , " + _
				"		par_value nvarchar (250) NULL , " + _
				"		par_sito_id int NOT NULL " + _
				" ); " + _
				" ALTER TABLE tb_siti_parametri WITH NOCHECK ADD CONSTRAINT PK_tb_siti_parametri PRIMARY KEY NONCLUSTERED (par_id); " + _
				" CREATE TABLE dbo.tb_admin ( " + _
				"		id_admin int IDENTITY (1, 1) NOT NULL , " + _
				"		admin_login nvarchar (50) NULL , " + _
				"		admin_password nvarchar (50) NULL , " + _
				"		admin_matricola nvarchar (50) NULL , " + _
				"		admin_nome nvarchar (50) NULL , " + _
				"		admin_cognome nvarchar (50) NULL , " + _
				"		admin_data_Nasc smalldatetime NULL , " + _
				"		admin_data_assunz smalldatetime NULL , " + _
				"		admin_note text NULL , " + _
				"		admin_ufficio int NULL , " + _
				"		admin_contratto int NULL , " + _
				"		admin_email nvarchar (100) NULL , " + _
				"		admin_direttore bit NULL , " + _
				"		admin_contatto int NULL , " + _
				"		admin_profilo int NULL , " + _
				"		admin_scadenza smalldatetime NULL " + _
				" ); " + _
				" ALTER TABLE tb_admin WITH NOCHECK ADD CONSTRAINT PK_tb_admin PRIMARY KEY NONCLUSTERED (id_admin); " + _
				" CREATE TABLE dbo.tb_Utenti ( " + _
				"		ut_ID int IDENTITY (1, 1) NOT NULL , " + _
				"		ut_NextCom_ID int NOT NULL , " + _
				" 		ut_login nvarchar (50) NULL , " + _
				"		ut_password nvarchar (50) NULL , " + _
				"		ut_Abilitato bit NULL , " + _
				"		ut_ScadenzaAccesso smalldatetime NULL " + _
				" ); " + _
				" ALTER TABLE tb_Utenti WITH NOCHECK ADD CONSTRAINT PK_tb_Utenti PRIMARY KEY NONCLUSTERED (ut_ID); " + _
				" CREATE TABLE dbo.rel_admin_sito ( " + _
				"		id_p int IDENTITY (1, 1) NOT NULL , " + _
				"		admin_id int NULL , " + _
				"		sito_id int NULL , " + _
				"		rel_as_permesso int NULL " + _
				" ); " + _
				" ALTER TABLE rel_admin_sito WITH NOCHECK ADD CONSTRAINT PK_rel_admin_sito PRIMARY KEY NONCLUSTERED (id_p); " + _
				" CREATE TABLE dbo.rel_utenti_sito ( " + _
				"		rel_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rel_ut_id int NOT NULL , " + _
				"		rel_sito_id int NOT NULL , " + _
				"		rel_permesso int NOT NULL " + _
				" ); " + _
				" ALTER TABLE rel_utenti_sito WITH NOCHECK ADD CONSTRAINT PK_rel_utenti_sito PRIMARY KEY NONCLUSTERED (rel_id); " + _
				" CREATE TABLE dbo.log_admin ( " + _
				"		log_id int IDENTITY (1, 1) NOT NULL , " + _
				"		log_admin_id int NULL , " + _
				"		log_sito_id int NULL , " + _
				"		log_data smalldatetime NULL , " + _
				"		log_username nvarchar (50) NULL " + _
				" ); " + _
				" ALTER TABLE log_admin WITH NOCHECK ADD CONSTRAINT PK_log_admin PRIMARY KEY CLUSTERED (log_id); " + _
				" CREATE TABLE dbo.log_utenti ( " + _
				"		log_id int IDENTITY (1, 1) NOT NULL , " + _
				"		log_ut_id int NOT NULL , " + _
				"		log_sito_id int NOT NULL , " + _
				"		log_data smalldatetime NULL , " + _
				"		log_username nvarchar (50) NULL " + _
				" ); " + _
				" ALTER TABLE log_utenti WITH NOCHECK ADD CONSTRAINT PK_log_utenti PRIMARY KEY NONCLUSTERED (log_id); " + _
				" ALTER TABLE log_admin ADD " + _
				"		CONSTRAINT FK_log_admin_tb_admin FOREIGN KEY (log_admin_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_log_admin_tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE log_utenti ADD " + _
				"		CONSTRAINT FK_log_utenti__tb_utenti FOREIGN KEY (log_ut_id) REFERENCES tb_Utenti (ut_ID) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_log_utenti__tb_siti FOREIGN KEY (log_sito_id) REFERENCES tb_siti (id_sito) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_admin_sito ADD " + _
				"		CONSTRAINT FK_rel_admin_sito__tb_admin FOREIGN KEY (admin_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_rel_admin_sito__tb_siti FOREIGN KEY (sito_id) REFERENCES tb_siti (id_sito) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_utenti_sito ADD " + _
				"		CONSTRAINT FK_rel_utenti_sito__tb_utenti FOREIGN KEY (rel_ut_id) REFERENCES tb_Utenti (ut_ID) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_rel_utenti_sito__tb_siti FOREIGN KEY (rel_sito_id) REFERENCES tb_siti (id_sito) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_siti_parametri ADD " + _
				"		CONSTRAINT FK_tb_siti_parametri_tb_siti FOREIGN KEY (par_sito_id) REFERENCES tb_siti (id_sito) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" INSERT INTO tb_admin(admin_nome, admin_cognome, admin_email, admin_login, admin_password) " + _
				" VALUES ('Supporto tecnico', 'Combinario', 'supporto@combinario.com', 'COMBINARIO', 'combitmp24'); " + _
				" INSERT INTO tb_admin(admin_nome, admin_cognome, admin_email, admin_login, admin_password) " + _
				" VALUES ('Spedizione email', 'SISTEMA', 'sviluppo@combinario.com', 'SISTEMA', '5315sis'); " + _
				" INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3) " + _
				" VALUES (1, 'NEXT-passport [gestione utenti]', 'NEXTpassport', 1, 'PASS_ADMIN', 'PASS_AMMINISTRATORI', 'PASS_UTENTI'); " + _
				" INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2, sito_p3) " + _
				" VALUES (3, 'NEXT-com [gestione comunicazioni]', 'NEXTcom', 1, 'COM_ADMIN', 'COM_USER', 'COM_POWER'); " + _
				" INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1, sito_p2) " + _
				" VALUES (25, 'NEXT-web 4.0 [gestione grafica e contenuti]', 'NEXTweb4', 1, 'WEB_ADMIN', 'WEB_USER'); " + _
				" INSERT INTO tb_siti_parametri(par_key, par_sito_id) VALUES('GruppoLavoroAreaRiservata', 1); " + _
				" INSERT INTO rel_admin_sito(admin_id, rel_as_permesso, sito_id) VALUES (1, 1, 3); " + _
				" INSERT INTO rel_admin_sito(admin_id, rel_as_permesso, sito_id) VALUES (1, 1, 1); " + _
				" INSERT INTO rel_admin_sito(admin_id, rel_as_permesso, sito_id) VALUES (1, 1, 25); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-COM
'...........................................................................................
'aggiunge il NEXT-COM. e' necessario installare precedentemente il NEXT-PASSPORT.
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTCOM(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTCOM = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTCOM = _
				" CREATE TABLE dbo.al_attivita_gruppi ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_gruppo_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_attivita_gruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_al_messaggi_gruppi PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_attivita_utenti ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_utente_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_attivita_utenti WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_al_attivita_utenti PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_default_gruppi ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_gruppo_id int NULL , " + _
				"		al_tipo_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_default_gruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_al_default_gruppi PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_default_utenti ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_utente_id int NULL , " + _
				"		al_tipo_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_default_utenti WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_al_default PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_documenti_gruppi ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_gruppo_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_documenti_gruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_documenti_gruppi PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_documenti_utenti ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_utente_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_documenti_utenti WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_documenti_utenti PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_pratiche_gruppi ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_gruppo_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_pratiche_gruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_pratica_gruppo PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.al_pratiche_utenti ( " + _
				"		al_id int IDENTITY (1, 1) NOT NULL , " + _
				"		al_tipo_id int NULL , " + _
				"		al_utente_id int NULL  " + _
				" ); " + _
				" ALTER TABLE al_pratiche_utenti WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_pratiche_utenti PRIMARY KEY CLUSTERED (al_id); " + _
				" CREATE TABLE dbo.log_cnt_email ( " + _
				"		log_id int IDENTITY (1, 1) NOT NULL , " + _
				"		log_cnt_id int NULL , " + _
				"		log_email nvarchar (250) NULL , " + _
				"		log_email_id int NULL  " + _
				" ); " + _
				" ALTER TABLE log_cnt_email WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_log_cnt_email PRIMARY KEY NONCLUSTERED (log_id); " + _
				" CREATE TABLE dbo.rel_dip_email ( " + _
				"		rel_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rel_emailSender nvarchar (250) NULL , " + _
				"		rel_emailSenderID int NULL , " + _
				"		rel_emailID int NULL , " + _
				"		rel_dipID int NULL , " + _
				"		rel_Read bit NOT NULL , " + _
				"		rel_Reply bit NOT NULL  " + _
				" ); " + _
				" ALTER TABLE rel_dip_email WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_dip_email PRIMARY KEY NONCLUSTERED (rel_id); " + _
				" CREATE TABLE dbo.rel_documenti_descrittori ( " + _
				"		rdd_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rdd_valore nvarchar (255) NULL , " + _
				"		rdd_documento_id int NULL , " + _
				"		rdd_descrittore_id int NULL  " + _
				" ); " + _
				" ALTER TABLE rel_documenti_descrittori WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_documenti_descrittori PRIMARY KEY CLUSTERED (rdd_id); " + _
				" CREATE TABLE dbo.rel_documenti_files ( " + _
				"		rel_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rel_documento_id int NOT NULL , " + _
				"		rel_files_id int NOT NULL  " + _
				" ); " + _
				" ALTER TABLE rel_documenti_files WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_documenti_files PRIMARY KEY CLUSTERED (rel_id); " + _
				" CREATE TABLE dbo.rel_rub_ind ( " + _
				"		id_rub_ind int IDENTITY (1, 1) NOT NULL , " + _
				"		id_indirizzo int NULL , " + _
				"		id_rubrica int NULL  " + _
				" ); " + _
				" ALTER TABLE rel_rub_ind WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_rub_ind PRIMARY KEY NONCLUSTERED (id_rub_ind); " + _
				" CREATE TABLE dbo.rel_tipologie_descrittori ( " + _
				"		rtd_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rtd_tipologia_id int NULL , " + _
				"		rtd_descrittore_id int NULL  " + _
				" ); " + _
				" ALTER TABLE rel_tipologie_descrittori WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_rel_tipologie_descrittori PRIMARY KEY CLUSTERED (rtd_id); " + _
				" CREATE TABLE dbo.tb_Files ( " + _
				"		F_id int IDENTITY (1, 1) NOT NULL , " + _
				"		F_original_name nvarchar (250) NULL , " + _
				"		F_encoded_name nvarchar (250) NULL , " + _
				"		F_size int NULL , " + _
				"		F_Data smalldatetime NULL , " + _
				"		F_allegato bit NULL , " + _
				"		F_original_path nvarchar (250) NULL , " + _
				"		F_encoded_path nvarchar (250) NULL , " + _
				"		F_LastUpdate smalldatetime NULL  " + _
				" ); " + _
				" ALTER TABLE tb_Files WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_Files PRIMARY KEY CLUSTERED (F_id); " + _
				" CREATE TABLE dbo.tb_Indirizzario ( " + _
				"		IDElencoIndirizzi int IDENTITY (1, 1) NOT NULL , " + _
				"		NomeElencoIndirizzi nvarchar (100) NULL , " + _
				"		SecondoNomeElencoIndirizzi nvarchar (30) NULL , " + _
				"		CognomeElencoIndirizzi nvarchar (100) NULL , " + _
				"		TitoloElencoIndirizzi nvarchar (50) NULL , " + _
				"		NomeOrganizzazioneElencoIndirizzi nvarchar (255) NULL , " + _
				"		QualificaElencoIndirizzi nvarchar (250) NULL , " + _
				"		IndirizzoElencoIndirizzi nvarchar (255) NULL , " + _
				"		CittaElencoIndirizzi nvarchar (50) NULL , " + _
				"		StatoProvElencoIndirizzi nvarchar (50) NULL , " + _
				"		ZonaElencoIndirizzi nvarchar (50) NULL , " + _
				"		CAPElencoIndirizzi nvarchar (20) NULL , " + _
				"		CountryElencoIndirizzi nvarchar (50) NULL , " + _
				"		DTNASCElencoIndirizzi smalldatetime NULL , " + _
				"		NoteElencoIndirizzi ntext NULL , " + _
				"		isSocieta bit NOT NULL , " + _
				"		ModoRegistra nvarchar (255) NULL , " + _
				"		DataIscrizione smalldatetime NULL , " + _
				"		LockedByApplication int NULL , " + _
				"		ApplicationsLocker nvarchar (50) NULL , " + _
				"		SyncroKey nvarchar (50) NULL , " + _
				"		SyncroTable nvarchar (50) NULL , " + _
				"		SyncroApplication int NULL , " + _
				"		LocalitaElencoIndirizzi nvarchar (100) NULL , " + _
				"		PraticaPrefisso nvarchar (5) NULL , " + _
				"		PraticaCount int NULL , " + _
				"		LuogoNascita nvarchar (255) NULL , " + _
				"		CF nvarchar (16) NULL , " + _
				"		cntRel int NULL , " + _
				"		lingua varchar (2) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_Indirizzario WITH NOCHECK ADD " + _
				"		CONSTRAINT DF__tb_indiri__Prati__017F0B4C DEFAULT (0) FOR PraticaCount, " + _
				"		CONSTRAINT PK_tb_Indirizzario PRIMARY KEY NONCLUSTERED (IDElencoIndirizzi); " + _
				" CREATE TABLE dbo.tb_ValoriNumeri ( " + _
				"		id_ValoreNumero int IDENTITY (1, 1) NOT NULL , " + _
				"		id_Indirizzario int NULL , " + _
				"		id_TipoNumero int NULL , " + _
				"		ValoreNumero nvarchar (250) NULL , " + _
				"		email_default bit NOT NULL , " + _
				"		SyncroField nvarchar (50) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_ValoriNumeri WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_ValoriNumeri PRIMARY KEY NONCLUSTERED (id_ValoreNumero); " + _
				" CREATE TABLE dbo.tb_allegati ( " + _
				"		all_id int IDENTITY (1, 1) NOT NULL , " + _
				"		all_attivita_id int NULL , " + _
				"		all_documento_id int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_allegati WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_allegati PRIMARY KEY CLUSTERED (all_id); " + _
				" CREATE TABLE dbo.tb_attivita ( " + _
				"		att_id int IDENTITY (1, 1) NOT NULL , " + _
				"		att_oggetto nvarchar (255) NULL , " + _
				"		att_testo ntext NULL , " + _
				"		att_note ntext NULL , " + _
				"		att_dataCrea smalldatetime NULL , " + _
				"		att_dataChiusa smalldatetime NULL , " + _
				"		att_dataS smalldatetime NULL , " + _
				"		att_priorita bit NULL , " + _
				"		att_conclusa bit NULL , " + _
				"		att_pubblica bit NULL , " + _
				"		att_eredita bit NULL , " + _
				"		att_sistema bit NULL , " + _
				"		att_domanda_id int NULL , " + _
				"		att_mittente_id int NULL , " + _
				"		att_pratica_id int NULL , " + _
				"		att_inSospeso bit NULL , " + _
				"		att_utente_chiusura int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_attivita WITH NOCHECK ADD " + _
				"		CONSTRAINT DF_tb_attivita_att_priorita DEFAULT (0) FOR att_priorita, " + _
				"		CONSTRAINT DF_tb_messaggi_msg_conclusa DEFAULT (0) FOR att_conclusa, " + _
				"		CONSTRAINT DF_tb_attivita_att_sistema DEFAULT (0) FOR att_sistema, " + _
				"		CONSTRAINT PK_tb_messaggi PRIMARY KEY CLUSTERED (att_id); " + _
				" CREATE TABLE dbo.tb_cnt_lingue ( " + _
				"		lingua_codice varchar (2) NOT NULL , " + _
				"		lingua_nome_IT varchar (20) NULL , " + _
				"		lingua_nome varchar (20) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_cnt_lingue WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_cnt_lingue PRIMARY KEY CLUSTERED (lingua_codice); " + _
				" CREATE TABLE dbo.tb_descrittori ( " + _
				"		descr_id int IDENTITY (1, 1) NOT NULL , " + _
				"		descr_nome nvarchar (50) NULL , " + _
				"		descr_tipo smallint NULL , " + _
				"		descr_ordine int NULL , " + _
				"		descr_principale bit NULL  " + _
				" ); " + _
				" ALTER TABLE tb_descrittori WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_descrittori PRIMARY KEY CLUSTERED (descr_id); " + _
				" CREATE TABLE dbo.tb_documenti ( " + _
				"		doc_id int IDENTITY (1, 1) NOT NULL , " + _
				"		doc_nome nvarchar (255) NULL , " + _
				"		doc_dataC smalldatetime NULL , " + _
				"		doc_pubblica bit NULL , " + _
				"		doc_eredita bit NULL , " + _
				"		doc_note ntext NULL , " + _
				"		doc_tipologia_id int NULL , " + _
				"		doc_pratica_id int NULL , " + _
				"		doc_creatore_id int NULL , " + _
				"		doc_mod_data smalldatetime NULL , " + _
				"		doc_mod_utente int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_documenti WITH NOCHECK ADD " + _
				"		CONSTRAINT DF_tb_documenti_doc_pubblica DEFAULT (0) FOR doc_pubblica, " + _
				"		CONSTRAINT DF_tb_documenti_doc_eredita DEFAULT (1) FOR doc_eredita, " + _
				"		CONSTRAINT PK_tb_documenti PRIMARY KEY CLUSTERED (doc_id); " + _
				" CREATE TABLE dbo.tb_email ( " + _
				"		email_id int IDENTITY (1, 1) NOT NULL , " + _
				"		email_text ntext NULL , " + _
				"		email_object nvarchar (200) NULL , " + _
				"		email_data smalldatetime NULL , " + _
				"		email_dipgenera int NULL , " + _
				"		email_docs ntext NULL , " + _
				"		email_page_ID int NULL , " + _
				"		email_page_owned bit NOT NULL , " + _
				"		email_in bit NOT NULL , " + _
				"		email_MessageID nvarchar (100) NULL , " + _
				"		email_UIDL nvarchar (250) NULL , " + _
				"		email_Account int NOT NULL , " + _
				"		email_To nvarchar (250) NULL , " + _
				"		email_CC nvarchar (250) NULL , " + _
				"		email_mime nvarchar (50) NULL , " + _
				"		email_From nvarchar (250) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_email WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_email PRIMARY KEY NONCLUSTERED (email_id); " + _
				" CREATE TABLE dbo.tb_emailConfig ( " + _
				"		config_id int IDENTITY (1, 1) NOT NULL , " + _
				"		config_host nvarchar (250) NULL , " + _
				"		config_port int NULL , " + _
				"		config_user nvarchar (50) NULL , " + _
				"		config_pass nvarchar (50) NULL , " + _
				"		config_protocol nvarchar (5) NULL , " + _
				"		config_email nvarchar (250) NULL , " + _
				"		config_deleteMessage bit NOT NULL , " + _
				"		config_delayDelMessage int NULL , " + _
				"		config_id_empl int NOT NULL  " + _
				" ); " + _
				" ALTER TABLE tb_emailConfig WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_emailConfig PRIMARY KEY NONCLUSTERED (config_id); " + _
				" CREATE TABLE dbo.tb_gruppi ( " + _
				"		id_Gruppo int IDENTITY (1, 1) NOT NULL , " + _
				"		nome_Gruppo nvarchar (50) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_gruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_gruppi PRIMARY KEY NONCLUSTERED (id_Gruppo); " + _
				" CREATE TABLE dbo.tb_pratiche ( " + _
				"		pra_id int IDENTITY (1, 1) NOT NULL , " + _
				"		pra_codice nvarchar (50) NULL , " + _
				"		pra_nome nvarchar (255) NULL , " + _
				"		pra_dataI smalldatetime NULL , " + _
				"		pra_dataUM smalldatetime NULL , " + _
				"		pra_dataA smalldatetime NULL , " + _
				"		pra_archiviata bit NULL , " + _
				"		pra_note ntext NULL , " + _
				"		pra_pubblica bit NULL , " + _
				"		pra_cliente_id int NULL , " + _
				"		pra_creatore_id int NULL , " + _
				"		pra_mod_data smalldatetime NULL , " + _
				"		pra_mod_utente int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_pratiche WITH NOCHECK ADD " + _
				"		CONSTRAINT DF_tb_pratiche_pra_archiviata DEFAULT (0) FOR pra_archiviata, " + _
				"		CONSTRAINT PK_tb_pratiche PRIMARY KEY CLUSTERED (pra_id); " + _
				" CREATE TABLE dbo.tb_rel_dipgruppi ( " + _
				"		id_rel_dipgruppi int IDENTITY (1, 1) NOT NULL , " + _
				"		id_impiegato int NULL , " + _
				"		id_gruppo int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_rel_dipgruppi WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_rel_dipgruppi PRIMARY KEY NONCLUSTERED (id_rel_dipgruppi); " + _
				" CREATE TABLE dbo.tb_rel_gruppirubriche ( " + _
				"		id_rel_grupprub int IDENTITY (1, 1) NOT NULL , " + _
				"		id_dellaRubrica int NULL , " + _
				"		id_Gruppo_assegnato int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_rel_gruppirubriche WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_rel_gruppirubriche PRIMARY KEY NONCLUSTERED (id_rel_grupprub); " + _
				" CREATE TABLE dbo.tb_rubriche ( " + _
				"		id_Rubrica int IDENTITY (1, 1) NOT NULL , " + _
				"		nome_Rubrica nvarchar (250) NULL , " + _
				"		note_Rubrica ntext NULL , " + _
				"		locked_rubrica bit NOT NULL , " + _
				"		rubrica_esterna bit NOT NULL , " + _
				"		SyncroTable nvarchar (50) NULL , " + _
				"		SyncroFilterTable nvarchar (50) NULL , " + _
				"		SyncroFilterKey int NULL  " + _
				" ); " + _
				" ALTER TABLE tb_rubriche WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_rubriche PRIMARY KEY NONCLUSTERED (id_Rubrica); " + _
				" CREATE TABLE dbo.tb_tipNumeri ( " + _
				"		id_tipoNumero int NOT NULL , " + _
				"		nome_tipoNumero nvarchar (250) NULL , " + _
				"		tipoNumero nvarchar (1) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_tipNumeri WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_tipNumeri PRIMARY KEY NONCLUSTERED (id_tipoNumero); " + _
				" CREATE TABLE dbo.tb_tipologie ( " + _
				"		tipo_id int IDENTITY (1, 1) NOT NULL , " + _
				"		tipo_nome nvarchar (50) NULL  " + _
				" ); " + _
				" ALTER TABLE tb_tipologie WITH NOCHECK ADD " + _
				"		CONSTRAINT PK_tb_tipologie PRIMARY KEY CLUSTERED (tipo_id); " + _
				" ALTER TABLE al_attivita_gruppi ADD " + _
				"		CONSTRAINT FK_al_messaggi_gruppi_tb_attivita FOREIGN KEY (al_tipo_id) REFERENCES tb_attivita (att_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_messaggi_gruppi_tb_gruppi FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_Gruppo) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_attivita_utenti ADD " + _
				"		CONSTRAINT FK_al_attivita_utenti_tb_admin FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_attivita_utenti_tb_attivita FOREIGN KEY (al_tipo_id) REFERENCES tb_attivita (att_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_default_gruppi ADD " + _
				"		CONSTRAINT FK_al_default_gruppi_tb_gruppi FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_Gruppo) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_default_gruppi_tb_pratiche FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_default_utenti ADD " + _
				"		CONSTRAINT FK_al_default_utenti_tb_admin FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_default_utenti_tb_pratiche FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_documenti_gruppi ADD " + _
				"		CONSTRAINT FK_al_documenti_gruppi_tb_documenti FOREIGN KEY (al_tipo_id) REFERENCES tb_documenti (doc_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_documenti_gruppi_tb_gruppi FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_Gruppo) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; "+ _
				" ALTER TABLE al_documenti_utenti ADD " + _
				"		CONSTRAINT FK_al_documenti_utenti_tb_admin FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_documenti_utenti_tb_documenti FOREIGN KEY (al_tipo_id) REFERENCES tb_documenti (doc_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_pratiche_gruppi ADD " + _
				"		CONSTRAINT FK_al_pratiche_gruppi_tb_gruppi FOREIGN KEY (al_gruppo_id) REFERENCES tb_gruppi (id_Gruppo) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_pratiche_gruppi_tb_pratiche FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE al_pratiche_utenti ADD " + _
				"		CONSTRAINT FK_al_pratiche_utenti_tb_admin FOREIGN KEY (al_utente_id) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_al_pratiche_utenti_tb_pratiche FOREIGN KEY (al_tipo_id) REFERENCES tb_pratiche (pra_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE log_cnt_email ADD " + _
				"		CONSTRAINT FK_log_cnt_email_tb_email FOREIGN KEY (log_email_id) REFERENCES tb_email (email_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_log_cnt_email_tb_Indirizzario FOREIGN KEY (log_cnt_id) " + _
				"		REFERENCES tb_Indirizzario (IDElencoIndirizzi) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_dip_email ADD " + _
				"		CONSTRAINT FK_rel_dip_email_tb_admin FOREIGN KEY (rel_dipID) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_rel_dip_email_tb_admin1 FOREIGN KEY (rel_emailSenderID) REFERENCES tb_admin (id_admin), " + _
				"		CONSTRAINT FK_rel_dip_email_tb_email FOREIGN KEY (rel_emailID) REFERENCES tb_email (email_id) " + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE; " + _
				" alter table rel_dip_email nocheck constraint FK_rel_dip_email_tb_admin1; " + _
				" ALTER TABLE rel_documenti_descrittori ADD " + _
				"		CONSTRAINT FK_rel_documenti_descrittori_tb_descrittori FOREIGN KEY (rdd_descrittore_id) " + _
				"		REFERENCES tb_descrittori (descr_id) ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_rel_documenti_descrittori_tb_documenti FOREIGN KEY (rdd_documento_id) " + _
				"		REFERENCES tb_documenti (doc_id) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_documenti_files ADD " + _
				"		CONSTRAINT FK_rel_documenti_files__tb_documenti FOREIGN KEY (rel_documento_id) " + _
				"		REFERENCES tb_documenti (doc_id) ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_rel_documenti_files__tb_files FOREIGN KEY (rel_files_id) REFERENCES tb_Files (F_id) " + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_rub_ind ADD " + _
				"		CONSTRAINT FK_rel_rub_ind_tb_Indirizzario FOREIGN KEY (id_indirizzo) " + _
				"		REFERENCES tb_Indirizzario (IDElencoIndirizzi) ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_rel_rub_ind_tb_rubriche FOREIGN KEY (id_rubrica) " + _
				"		REFERENCES tb_rubriche (id_rubrica) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE rel_tipologie_descrittori ADD " + _
				"		CONSTRAINT FK_rel_tipologie_descrittori_tb_descrittori FOREIGN KEY (rtd_descrittore_id) " + _
				"		REFERENCES tb_descrittori (descr_id) ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_rel_tipologie_descrittori_tb_tipologie FOREIGN KEY (rtd_tipologia_id) " + _
				"		REFERENCES tb_tipologie (tipo_id) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_Indirizzario ADD " + _
				"		CONSTRAINT FK_tb_indirizzario__tb_cnt_lingue FOREIGN KEY (lingua) " + _
				"		REFERENCES tb_cnt_lingue (lingua_codice) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_ValoriNumeri ADD " + _
				"		CONSTRAINT FK_tb_ValoriNumeri_tb_Indirizzario FOREIGN KEY (id_Indirizzario) " + _
				"		REFERENCES tb_Indirizzario (IDElencoIndirizzi) ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_tb_ValoriNumeri_tb_tipNumeri FOREIGN KEY (id_TipoNumero) " + _
				"		REFERENCES tb_tipNumeri (id_tipoNumero) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_allegati ADD " + _
				"		CONSTRAINT FK_tb_allegati__tb_documenti FOREIGN KEY (all_documento_id) " + _
				"		REFERENCES tb_documenti (doc_id) ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_tb_allegati_tb_messaggi FOREIGN KEY (all_attivita_id) REFERENCES tb_attivita (att_id) " + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_attivita ADD " + _
				"		CONSTRAINT FK_tb_attivita_tb_admin FOREIGN KEY (att_mittente_id) REFERENCES tb_admin (id_admin), " + _
				"		CONSTRAINT FK_tb_attivita_tb_pratiche FOREIGN KEY (att_pratica_id) REFERENCES tb_pratiche (pra_id); " + _
				" alter table tb_attivita nocheck constraint FK_tb_attivita_tb_admin; " + _
				" alter table tb_attivita nocheck constraint FK_tb_attivita_tb_pratiche; " + _
				" ALTER TABLE tb_documenti ADD " + _
				"		CONSTRAINT FK_tb_documenti_tb_admin FOREIGN KEY (doc_creatore_id) REFERENCES tb_admin (id_admin), " + _
				"		CONSTRAINT FK_tb_documenti_tb_pratiche FOREIGN KEY (doc_pratica_id) REFERENCES tb_pratiche (pra_id), " + _
				"		CONSTRAINT FK_tb_documenti_tb_tipologie FOREIGN KEY (doc_tipologia_id) " + _
				"		REFERENCES tb_tipologie (tipo_id) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" alter table tb_documenti nocheck constraint FK_tb_documenti_tb_admin; " + _
				" alter table tb_documenti nocheck constraint FK_tb_documenti_tb_pratiche; " + _
				" ALTER TABLE tb_emailConfig ADD " + _
				"		CONSTRAINT FK_tb_emailConfig_tb_admin FOREIGN KEY (config_id_empl) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_pratiche ADD " + _
				"		CONSTRAINT FK_tb_pratiche_tb_admin FOREIGN KEY (pra_creatore_id) REFERENCES tb_admin (id_admin), " + _
				"		CONSTRAINT FK_tb_pratiche_tb_Indirizzario FOREIGN KEY (pra_cliente_id) " + _
				"		REFERENCES tb_Indirizzario (IDElencoIndirizzi) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" alter table tb_pratiche nocheck constraint FK_tb_pratiche_tb_admin; " + _
				" ALTER TABLE tb_rel_dipgruppi ADD " + _
				"		CONSTRAINT FK_tb_rel_dipgruppi_tb_admin FOREIGN KEY (id_impiegato) REFERENCES tb_admin (id_admin) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE , " + _
				"		CONSTRAINT FK_tb_rel_dipgruppi_tb_gruppi FOREIGN KEY (id_gruppo) REFERENCES tb_gruppi (id_Gruppo) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_rel_gruppirubriche ADD " + _
				"		CONSTRAINT FK_tb_rel_gruppirubriche_tb_gruppi FOREIGN KEY (id_Gruppo_assegnato) " + _
				"		REFERENCES tb_gruppi (id_Gruppo) ON DELETE CASCADE ON UPDATE CASCADE, " + _
				"		CONSTRAINT FK_tb_rel_gruppirubriche_tb_rubriche FOREIGN KEY (id_dellaRubrica) " + _
				"		REFERENCES tb_rubriche (id_rubrica) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" ALTER TABLE tb_Utenti ADD " + _
				"		CONSTRAINT FK_tb_Utenti_tb_Indirizzario FOREIGN KEY (ut_NextCom_ID) " + _
				"		REFERENCES tb_Indirizzario (IDElencoIndirizzi) ON DELETE CASCADE ON UPDATE CASCADE; " + _
				" INSERT INTO tb_cnt_lingue(lingua_codice, lingua_nome_it, lingua_nome) VALUES ('de', 'Tedesco', 'Deutsch'); " + _
				" INSERT INTO tb_cnt_lingue(lingua_codice, lingua_nome_it, lingua_nome) VALUES ('en', 'Inglese', 'English'); " + _
				" INSERT INTO tb_cnt_lingue(lingua_codice, lingua_nome_it, lingua_nome) VALUES ('es', 'Spagnolo', 'Espanol'); " + _
				" INSERT INTO tb_cnt_lingue(lingua_codice, lingua_nome_it, lingua_nome) VALUES ('fr', 'Francese', 'Fransais'); " + _
				" INSERT INTO tb_cnt_lingue(lingua_codice, lingua_nome_it, lingua_nome) VALUES ('it', 'Italiano', 'Italiano'); " + _
				" INSERT INTO tb_rubriche(nome_rubrica, locked_rubrica, rubrica_esterna) " + _
				" VALUES ('Supporto tecnico', 1, 0); " + _
				" INSERT INTO tb_rubriche(nome_rubrica, locked_rubrica, rubrica_esterna) " + _
				" VALUES ('Sito - Contatti', 1, 0); " + _
				" INSERT INTO tb_indirizzario(NomeOrganizzazioneElencoIndirizzi, IndirizzoElencoIndirizzi, CittaElencoIndirizzi, " + _
				"							  StatoProvElencoIndirizzi, CAPElencoIndirizzi, CountryElencoIndirizzi, isSocieta, " + _
				"							  ModoRegistra, PraticaCount, CF, lingua) " + _
				" VALUES ('Combinario', 'Via Giordano Bruno, 29', 'Mestre', 'Ve', '30174', 'Italia', 1, 'Combinario', 0, '04189660279', 'it'); " + _
				" INSERT INTO rel_rub_ind(id_indirizzo, id_rubrica) VALUES (1, 1); " + _
				" INSERT INTO tb_gruppi(nome_gruppo) VALUES ('Amministrazione'); " + _
				" INSERT INTO tb_rel_gruppiRubriche(id_dellaRubrica, id_gruppo_assegnato) VALUES (1, 1); " + _
				" INSERT INTO tb_rel_gruppiRubriche(id_dellaRubrica, id_gruppo_assegnato) VALUES (2, 1); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (1, 'Telefono'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (2, 'Telefono Ufficio'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (3, 'Telefono Cellulare'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (4, 'Telefono Casa'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (5, 'Numero Fax'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (6, 'Indirizzo Email'); " + _
				" INSERT INTO tb_tipNumeri(id_tipoNumero, nome_tipoNumero) VALUES (7, 'Indirizzo Web'); " + _
				" INSERT INTO tb_valoriNumeri(id_indirizzario, id_tipoNumero, valoreNumero, email_default) " + _
				" VALUES (1, 1, '041 8877149', 0); " + _
				" INSERT INTO tb_valoriNumeri(id_indirizzario, id_tipoNumero, valoreNumero, email_default) " + _
				" VALUES (1, 5, '041 8871249', 0); " + _
				" INSERT INTO tb_valoriNumeri(id_indirizzario, id_tipoNumero, valoreNumero, email_default) " + _
				" VALUES (1, 6, 'supporto@combinario.com', 1); " + _
				" INSERT INTO tb_valoriNumeri(id_indirizzario, id_tipoNumero, valoreNumero, email_default) " + _
				" VALUES (1, 7, 'www.combinario.com', 0); " + _
				" INSERT INTO tb_rel_dipGruppi(id_impiegato, id_gruppo) VALUES (1, 1); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-LINK
'...........................................................................................
'aggiunge le tabelle e relazioni per l'applicativo NEXT-LINK
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTLINK(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTLINK = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTLINK = _
				" CREATE TABLE dbo.tb_links_categorie ( " + _
			  	"		cat_id int IDENTITY (1, 1) NOT NULL , " + _
				"		cat_nome_IT nvarchar (50) NULL , " + _
			  	"		cat_nome_EN nvarchar (50) NULL , " + _
			  	"		cat_nome_FR nvarchar (50) NULL , " + _
			  	"		cat_nome_DE nvarchar (50) NULL , " + _
			  	"		cat_nome_ES nvarchar (50) NULL , " + _
			  	"		CONSTRAINT PK_tb_links_categorie PRIMARY KEY CLUSTERED (cat_id) " + _
			  	"		); " + _
			  	" CREATE TABLE dbo.tb_links ( " + _
			  	"		link_id int IDENTITY (1, 1) NOT NULL , " + _
			  	"		link_cat_id int NULL , " + _
			  	"		link_nome_IT nvarchar (255) NULL , " + _
			  	"		link_nome_EN nvarchar (255) NULL , " + _
			  	"		link_nome_FR nvarchar (255) NULL , " + _
			  	"		link_nome_DE nvarchar (255) NULL , " + _
			  	"		link_nome_ES nvarchar (255) NULL , " + _
			  	"		link_url nvarchar (250) NULL , " + _
			  	"		link_logo nvarchar (50) NULL , " + _
			  	"		link_data_reset smalldatetime NULL , " + _
			  	"		link_count int NULL , " + _
			  	"		link_descr_IT ntext NULL , " + _
			  	"		link_descr_EN ntext NULL , " + _
			  	"		link_descr_FR ntext NULL , " + _
			  	"		link_descr_DE ntext NULL , " + _
			  	"		link_descr_ES ntext NULL , " + _
			  	"		link_ordine int NULL , " + _
			  	"		CONSTRAINT PK_tb_links PRIMARY KEY CLUSTERED (link_id) , " + _
			  	"		CONSTRAINT FK_tb_links_tb_links_categorie FOREIGN KEY (link_cat_id)  " + _
			  	"		REFERENCES tb_links_categorie (cat_id) ON DELETE NO ACTION  ON UPDATE NO ACTION" + _
			  	"		); " + _
				" ALTER TABLE tb_links NOCHECK CONSTRAINT FK_tb_links_tb_links_categorie ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-NEWS
'...........................................................................................
'aggiunge le tabelle e relazioni per l'applicativo NEXT-NEWS
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTNEWS(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTNEWS = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTNEWS = _
				" CREATE TABLE dbo.tb_news_categorie ( " + _
				"		cat_id int IDENTITY (1, 1) NOT NULL , " + _
				" 		cat_nome_IT nvarchar(50) NULL , " + _
				"		cat_nome_EN nvarchar(50) NULL , " + _
				"		cat_nome_FR nvarchar(50) NULL , " + _
				"		cat_nome_DE nvarchar(50) NULL , " + _
				"		cat_nome_ES nvarchar(50) NULL , " + _
				"		CONSTRAINT PK_tb_news_categorie PRIMARY KEY CLUSTERED (cat_id) " + _
				"		); " +_
				" CREATE TABLE dbo.tb_news ( " +_
				"		news_id int IDENTITY (1, 1) NOT NULL , " + _
				"		news_cat_id int NULL , " + _
				"		news_titolo_it nvarchar(255) NULL , " + _
				"		news_titolo_en nvarchar(255) NULL , " + _
				"		news_titolo_es nvarchar(255) NULL , " + _
				"		news_titolo_de nvarchar(255) NULL , " + _
				"		news_titolo_fr nvarchar(255) NULL , " + _
				"		news_dataPubbl smalldatetime NULL , " + _
				"		news_dataScad smalldatetime NULL , " + _
				"		news_img nvarchar(50) NULL , " + _
				"		news_pagina int NULL , " + _
				"		news_estratto_IT ntext NULL , " + _
				"		news_estratto_EN ntext NULL , " + _
				"		news_estratto_ES ntext NULL , " + _
				"		news_estratto_DE ntext NULL , " + _
				"		news_estratto_FR ntext NULL , " + _
				"		news_elenco bit NULL , " + _
				"		news_rotazione bit NULL , " + _
				"		news_url nvarchar(250) NULL , " + _
				"		CONSTRAINT PK_tb_news PRIMARY KEY CLUSTERED (news_id) , " + _
				"		CONSTRAINT FK_tb_news_tb_news_categorie FOREIGN KEY (news_cat_id) " + _
				"			REFERENCES tb_news_categorie (cat_id) ON DELETE NO ACTION ON UPDATE NO ACTION " + _
				"		);" + _
				" ALTER TABLE tb_news NOCHECK CONSTRAINT FK_tb_news_tb_news_categorie ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-GALLERY
'...........................................................................................
'aggiunge le tabelle e relazioni per l'applicativo NEXT-GALLERY
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTGALLERY(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTGALLERY = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTGALLERY = _
				" CREATE TABLE dbo.ptb_categorieGallery ( " + _
			  	"		catC_id int IDENTITY (1, 1) NOT NULL , " + _
				"		catC_nome_it nvarchar(250) NULL , " + _
				"		catC_nome_en nvarchar(250) NULL , " + _
				"		catC_nome_fr nvarchar(250) NULL , " + _
				"		catC_nome_de nvarchar(250) NULL , " + _
				"		catC_nome_es nvarchar(250) NULL , " + _
				"		catC_codice nvarchar(50) NULL , " + _
				"		catC_foglia bit NULL , " + _
				"		catC_livello int NULL , " + _
				"		catC_padre_id int NULL , " + _
				"		catC_tipologia_padre_base int NULL , " + _
				"		catC_ordine int NULL , " + _
				"		catC_ordine_assoluto nvarchar(250) NULL , " + _
				"		catC_descr_IT ntext NULL , " + _
				"		catC_descr_EN ntext NULL , " + _
				"		catC_descr_ES ntext NULL , " + _
				"		catC_descr_DE ntext NULL , " + _
				"		catC_descr_FR ntext NULL , " + _
				"		catC_visibile bit NULL , " + _
				"		catC_albero_visibile bit NULL " + _
				"		CONSTRAINT PK_ptb_categorieGallery PRIMARY KEY CLUSTERED (catC_id) " + _
				"		); " + _
				" CREATE TABLE dbo.ptb_Immagini ( " + _
				"		I_Id int IDENTITY (1, 1) NOT NULL , " + _
				"		I_Gallery_id int NOT NULL , " + _
				"		I_Titolo nvarchar(250) NULL , " + _
				"		I_Didascalia_IT ntext NULL , " + _
				"		I_Didascalia_EN ntext NULL , " + _
				"		I_Didascalia_ES ntext NULL , " + _
				"		I_Didascalia_DE ntext NULL , " + _
				"		I_Didascalia_FR ntext NULL , " + _
				"		I_Pubblicazione smalldatetime NULL , " + _
				"		I_Visibile bit NULL , " + _
				"		I_numero int NULL , " + _
				"		I_ordine int NULL , " + _
				"		I_thumb nvarchar(250) NULL , " + _
				"		I_zoom nvarchar(250) NULL , " + _
				"		CONSTRAINT PK_ptb_Immagini PRIMARY KEY CLUSTERED (I_Id) " + _
				"		); " + _
				" CREATE TABLE dbo.ptb_descrittori ( " + _
				"		des_ID int IDENTITY (1, 1) NOT NULL , " + _
				"		des_nome_it nvarchar(250) NULL , " + _
				"		des_nome_en nvarchar(250) NULL , " + _
				"		des_nome_fr nvarchar(250) NULL , " + _
				"		des_nome_de nvarchar(250) NULL , " + _
				"		des_nome_es nvarchar(250) NULL , " + _
				"		des_principale bit NULL , " + _
				"		des_tipo int NULL , " + _
				"		des_ordine int NULL , " + _
				"		CONSTRAINT PK_ptb_descrittori PRIMARY KEY CLUSTERED (des_ID) " + _
				"		); " + _
				" CREATE TABLE dbo.prel_descrittori_gallery ( " + _
				"		rdi_ID int IDENTITY (1, 1) NOT NULL , " + _
				"		rdi_descrittore_id int NOT NULL , " + _
				"		rdi_gallery_id int NOT NULL , " + _
				"		rdi_valore_it nvarchar(250) NULL , " + _
				"		rdi_valore_en nvarchar(250) NULL , " + _
				"		rdi_valore_fr nvarchar(250) NULL , " + _
				"		rdi_valore_de nvarchar(250) NULL , " + _
				"		rdi_valore_es nvarchar(250) NULL , " + _
				"		rdi_memo_it ntext NULL , " + _
				"		rdi_memo_en ntext NULL , " + _
				"		rdi_memo_fr ntext NULL , " + _
				"		rdi_memo_de ntext NULL , " + _
				"		rdi_memo_es ntext NULL , " + _
				"		CONSTRAINT PK_prel_descrittori_gallery PRIMARY KEY CLUSTERED (rdi_ID) " + _
				"		); " + _
				" CREATE TABLE dbo.ptb_gallery ( " + _
				"		gallery_id int IDENTITY (1, 1) NOT NULL , " + _
				"		gallery_name_it nvarchar(250) NULL , " + _
				"		gallery_name_en nvarchar(250) NULL , " + _
				"		gallery_name_fr nvarchar(250) NULL , " + _
				"		gallery_name_de nvarchar(250) NULL , " + _
				"		gallery_name_es nvarchar(250) NULL , " + _
				"		gallery_no int NULL , " + _
				"		gallery_idcategoria int NULL , " + _
				"		gallery_ordine int NULL , " + _
				"		gallery_visibile bit NULL , " + _
				"		gallery_codice nvarchar(250) NULL , " + _
				"		CONSTRAINT PK_gallery_id PRIMARY KEY CLUSTERED (gallery_id) " + _
				"		); " + _
				" ALTER TABLE ptb_gallery ADD CONSTRAINT FK_ptb_gallery__ptb_categorieGallery " + _
				"		FOREIGN KEY (gallery_idcategoria) REFERENCES ptb_categorieGallery (catC_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE ; " + _
				" ALTER TABLE ptb_Immagini ADD CONSTRAINT FK_ptb_Immagini__ptb_gallery " + _
				"		FOREIGN KEY (I_Gallery_id) REFERENCES ptb_gallery (gallery_ID) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE ; " + _
				" ALTER TABLE prel_descrittori_gallery ADD CONSTRAINT FK_prel_descrittori_gallery__ptb_descrittori " + _
				"		FOREIGN KEY (rdi_descrittore_id) REFERENCES ptb_descrittori (des_ID) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE ; " + _
				" ALTER TABLE prel_descrittori_gallery ADD CONSTRAINT FK_prel_descrittori_gallery__ptb_gallery " + _
				"		FOREIGN KEY (rdi_gallery_id) REFERENCES ptb_gallery (gallery_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-FAQ
'...........................................................................................
'aggiunge le tabelle e relazioni per l'applicativo NEXT-FAQ
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTFAQ(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTFAQ = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTFAQ = _
				" CREATE TABLE dbo.tb_FAQ_categorie (" + _
				" 		cat_id int  IDENTITY (1, 1) NOT NULL , " + _
				"		cat_nome_IT nvarchar(250) NULL , " + _
				"		cat_nome_EN nvarchar(250) NULL , " + _
				"		cat_nome_DE nvarchar(250) NULL , " + _
				"		cat_nome_FR nvarchar(250) NULL , " + _
				"		cat_nome_ES nvarchar(250) NULL , " + _
				"		cat_ordine int NULL , " + _
				"		CONSTRAINT PK_tb_FAQ_categorie PRIMARY KEY CLUSTERED (cat_id) " + _
				"		); " + _
				" CREATE TABLE dbo.tb_FAQ( " + _
				"		faq_id int IDENTITY (1, 1) NOT NULL , " + _
				"		faq_cat_id int NOT NULL , " + _
				"		faq_domanda_IT nvarchar(250) NULL , " + _
				"		faq_domanda_EN nvarchar(250) NULL , " + _
				"		faq_domanda_DE nvarchar(250) NULL , " + _
				"		faq_domanda_FR nvarchar(250) NULL , " + _
				"		faq_domanda_ES nvarchar(250) NULL , " + _
				"		faq_visibile bit NULL , " + _
				"		faq_ordine int NULL , " + _
				"		faq_risposta_IT ntext NULL , " + _
				"		faq_risposta_EN ntext NULL , " + _
				"		faq_risposta_DE ntext NULL , " + _
				"		faq_risposta_FR ntext NULL , " + _
				"		faq_risposta_ES ntext NULL , " + _
				"		CONSTRAINT PK_tb_FAQ PRIMARY KEY CLUSTERED (faq_id), " + _
				"		CONSTRAINT FK_tb_FAQ__tb_FAQ_categorie FOREIGN KEY (faq_cat_id) " + _
				"			REFERENCES tb_FAQ_categorie(cat_id) ON DELETE NO ACTION  ON UPDATE NO ACTION " + _
				"		);" + _
				" ALTER TABLE tb_FAQ NOCHECK CONSTRAINT FK_tb_FAQ__tb_FAQ_categorie ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-TEAM
'...........................................................................................
'aggiunge le tabelle e relazioni per l'applicativo NEXT-TEAM
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTTEAM(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTTEAM = ""
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTTEAM = _
				" CREATE TABLE dbo.Otb_livelli (" + _
				" 		lvl_id int  IDENTITY (1, 1) NOT NULL , " + _
				"		lvl_nome_IT nvarchar(250) NULL , " + _
				"		lvl_nome_EN nvarchar(250) NULL , " + _
				"		lvl_nome_DE nvarchar(250) NULL , " + _
				"		lvl_nome_FR nvarchar(250) NULL , " + _
				"		lvl_nome_ES nvarchar(250) NULL , " + _
				"		lvl_ordine int NULL , " + _
				"		CONSTRAINT PK_otb_livelli PRIMARY KEY CLUSTERED (lvl_id) " + _
				"		); " + _
				" CREATE TABLE dbo.Otb_componenti( " + _
				"		com_id int IDENTITY (1, 1) NOT NULL , " + _
				"		com_NEXTCOM_id int NOT NULL , " + _
				"		com_lvl_id int NOT NULL , " + _
				"		com_visibile BIT NULL , " + _
				"		com_ordine int NULL , " + _
				"		com_foto nvarchar(250) NULL , " + _
				"		com_posizione_IT nvarchar(250) NULL , " + _
				"		com_posizione_EN nvarchar(250) NULL , " + _
				"		com_posizione_DE nvarchar(250) NULL , " + _
				"		com_posizione_FR nvarchar(250) NULL , " + _
				"		com_posizione_ES nvarchar(250) NULL , " + _
				"		com_curriculum_IT ntext NULL , " + _
				"		com_curriculum_EN ntext NULL , " + _
				"		com_curriculum_DE ntext NULL , " + _
				"		com_curriculum_FR ntext NULL , " + _
				"		com_curriculum_ES ntext NULL , " + _
				"		CONSTRAINT PK_otb_componenti PRIMARY KEY CLUSTERED (com_id) " + _
				"		); " + _
				" ALTER TABLE Otb_componenti ADD CONSTRAINT FK_Otb_componenti__Otb_livelli " + _
				"		FOREIGN KEY (com_lvl_id) REFERENCES Otb_livelli (lvl_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE Otb_componenti ADD CONSTRAINT FK_Otb_componenti__tb_indirizzario " + _
				"		FOREIGN KEY (com_NEXTCOM_id) REFERENCES tb_indirizzario(IDElencoIndirizzi) " +_
				"		ON UPDATE CASCADE ON DELETE CASCADE; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'INSTALLA NEXT-WEB 5.0
'...........................................................................................
'aggiunge database per gestione next-web 5.0 (ex dbLayers)
'...........................................................................................
function Install__FRAMEWORK_CORE__NEXTWEB5(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__FRAMEWORK_CORE__NEXTWEB5 = _
				" CREATE TABLE tb_tipo ( " + _
				"	id_tip int NOT NULL , " + _
				"	tipo TEXT(50) WITH COMPRESSION NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_css_groups ( " + _
				"	grp_id int IDENTITY (1, 1) NOT NULL , " + _
				"	grp_id_webs int NOT NULL , " + _
				"	grp_name TEXT(250) WITH COMPRESSION NOT NULL , " + _
				"	grp_name_class TEXT(250) WITH COMPRESSION NOT NULL , " + _
				"	grp_for_editor bit NOT NULL , " + _
				"	grp_for_file bit NOT NULL , " + _
				"	grp_checksum TEXT(250) WITH COMPRESSION NOT NULL , " + _
				"	grp_ins_data smalldatetime NOT NULL , " + _
				"	grp_ins_admin_id int NOT NULL , " + _
				"	grp_mod_data smalldatetime NOT NULL , " + _
				"	grp_mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_css_styles ( " + _
				"	style_id int IDENTITY (1, 1) NOT NULL , " + _
				"	style_grp_id int NOT NULL , " + _
				"	style_class TEXT(20) WITH COMPRESSION NOT NULL , " + _
				"	style_pseudoclass TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_font_family TEXT(250) WITH COMPRESSION NULL , " + _
				"	style_font_size float NULL , " + _
				"	style_font_weight TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_font_style TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_color TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_background_color TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_line_height float NULL , " + _
				"	style_letter_spacing float NULL , " + _
				"	style_text_align TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_text_decoration TEXT(20) WITH COMPRESSION NULL , " + _
				"	style_description TEXT(250) WITH COMPRESSION NULL , " + _
				"	style_ins_data smalldatetime NOT NULL , " + _
				"	style_ins_admin_id int NOT NULL , " + _
				"	style_mod_data smalldatetime NOT NULL , " + _
				"	style_mod_admin_id smalldatetime NOT NULL " + _
				"	); " + _
				" CREATE TABLE tb_objects ( " + _
				"	id_objects int IDENTITY (1, 1) NOT NULL , " + _
				"	id_webs int NOT NULL , " + _
				"	name_objects TEXT(255) WITH COMPRESSION NOT NULL , " + _
				"	identif_objects TEXT(70) WITH COMPRESSION NOT NULL , " + _
				"	param_list TEXT WITH COMPRESSION NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_pages ( " + _
				"	id_page int IDENTITY (1, 1) NOT NULL , " + _
				"	id_webs int NOT NULL , " + _
				"	id_template int NULL , " + _
				"	lingua TEXT(2) WITH COMPRESSION NULL , " + _
				"	nomepage TEXT(250) WITH COMPRESSION NOT NULL , " + _
				"	template bit NOT NULL , " + _
				"	SfondoColore TEXT(7) WITH COMPRESSION NULL , " + _
				"	SfondoImmagine TEXT(255) WITH COMPRESSION NULL , " + _
				"	Contatore int NOT NULL , " + _
				"	ContRes smalldatetime NOT NULL , " + _
				"	contUtenti int NOT NULL , " + _
				"	contCrawler int NOT NULL , " + _
				"	contAltro int NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_pagineSito ( " + _
				"	id_pagineSito int IDENTITY (1, 1) NOT NULL , " + _
				"	id_web int NOT NULL , " + _
				"	archiviata bit NOT NULL , " + _
				"	riservata bit NOT NULL , " + _
				"	id_pagDyn_IT int NOT NULL , " + _
				"	id_pagDyn_EN int NULL , " + _
				"	id_pagDyn_FR int NULL , " + _
				"	id_pagDyn_DE int NULL , " + _
				"	id_pagDyn_ES int NULL , " + _
				"	id_pagStage_IT int NOT NULL , " + _
				"	id_pagStage_EN int NULL , " + _
				"	id_pagStage_FR int NULL , " + _
				"	id_pagStage_DE int NULL , " + _
				"	id_pagStage_ES int NULL , " + _
				"	nome_ps_IT TEXT(255) WITH COMPRESSION NOT NULL , " + _
				"	nome_ps_EN TEXT(255) WITH COMPRESSION NULL , " + _
				"	nome_ps_FR TEXT(255) WITH COMPRESSION NULL , " + _
				"	nome_ps_DE TEXT(255) WITH COMPRESSION NULL , " + _
				"	nome_ps_ES TEXT(255) WITH COMPRESSION NULL , " + _
				"	PAGE_keywords_IT TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_keywords_EN TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_keywords_FR TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_keywords_DE TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_keywords_ES TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_description_IT TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_description_EN TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_description_FR TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_description_DE TEXT WITH COMPRESSION NULL , " + _
				"	PAGE_description_ES TEXT WITH COMPRESSION NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_storico_pages ( " + _
				"	sp_ID int IDENTITY (1, 1) NOT NULL , " + _
				"	sp_sw_id int NOT NULL, " + _
				"	sp_page_id int NOT NULL , " + _
				"	sp_pagineSito_id int NOT NULL , " + _
				"	sp_nomepage TEXT(250) WITH COMPRESSION NOT NULL , " + _
				"	sp_lingua TEXT(2) WITH COMPRESSION NOT NULL , " + _
				"	sp_contatore int NOT NULL , " + _
				"	sp_contUtenti int NOT NULL , " + _
				"	sp_contCrawler int NOT NULL , " + _
				"	sp_contAltro int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_storico_webs ( " + _
				"	sw_ID int IDENTITY (1, 1) NOT NULL , " + _
				"	sw_webs_id int NOT NULL , " + _
				"	sw_data smalldatetime NOT NULL , " + _
				"	sw_contatore int NOT NULL , " + _
				"	sw_contUtenti int NOT NULL , " + _
				"	sw_contCrawler int NOT NULL , " + _
				"	sw_contAltro int NOT NULL , " + _
				"	sw_ins_data smalldatetime NOT NULL , " + _
				"	sw_ins_admin_id smalldatetime NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_webs ( " + _
				"	id_webs int NOT NULL , " + _
				"	nome_webs TEXT(50) WITH COMPRESSION NOT NULL , " + _
				"	lingua_iniziale TEXT(2) WITH COMPRESSION NOT NULL , " + _
				"	id_home_page int NULL , " + _
				"	id_home_page_riservata int NULL , " + _
				"	sito_in_aggiornamento bit NOT NULL , " + _
				"	sito_in_aggiornamento_pagina int NULL , " + _
				"	sito_in_costruzione bit NOT NULL , " + _
				"	sito_in_costruzione_pagina int NULL , " + _
				"	errore_pagina int NULL , " + _
				"	lingua_EN bit NOT NULL , " + _
				"	lingua_FR bit NOT NULL , " + _
				"	lingua_DE bit NOT NULL , " + _
				"	lingua_ES bit NOT NULL , " + _
				"	titolo_IT TEXT(255) WITH COMPRESSION NOT NULL , " + _
				"	titolo_EN TEXT(255) WITH COMPRESSION NULL , " + _
				"	titolo_FR TEXT(255) WITH COMPRESSION NULL , " + _
				"	titolo_DE TEXT(255) WITH COMPRESSION NULL , " + _
				"	titolo_ES TEXT(255) WITH COMPRESSION NULL , " + _
				"	META_Author TEXT WITH COMPRESSION NULL , " + _
				"	META_keywords_IT TEXT WITH COMPRESSION NULL , " + _
				"	META_keywords_EN TEXT WITH COMPRESSION NULL , " + _
				"	META_keywords_FR TEXT WITH COMPRESSION NULL , " + _
				"	META_keywords_DE TEXT WITH COMPRESSION NULL , " + _
				"	META_keywords_ES TEXT WITH COMPRESSION NULL , " + _
				"	META_description_IT TEXT WITH COMPRESSION NULL , " + _
				"	META_description_EN TEXT WITH COMPRESSION NULL , " + _
				"	META_description_FR TEXT WITH COMPRESSION NULL , " + _
				"	META_description_DE TEXT WITH COMPRESSION NULL , " + _
				"	META_description_ES TEXT WITH COMPRESSION NULL , " + _
				"	contatore int NOT NULL , " + _
				"	contRes smalldatetime NOT NULL , " + _
				"	contUtenti int NOT NULL , " + _
				"	contCrawler int NOT NULL , " + _
				"	contAltro int NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_layers (" + _
				"	id_lay int IDENTITY (1, 1) NOT NULL ," + _
				"	id_pag int NOT NULL ," + _
				"	id_tipo int NOT NULL ," + _
				"	id_objects int NULL ," + _
				"	tipo_contenuto TEXT(1) WITH COMPRESSION NOT NULL ," + _
				"	z_order int NOT NULL ," + _
				"	nome TEXT(250) WITH COMPRESSION NULL ," + _
				"	visibile bit NOT NULL ," + _
				"	x int NOT NULL ," + _
				"	y int NOT NULL ," + _
				"	largo int NOT NULL ," + _
				"	alto int NOT NULL ," + _
				"	em_x TEXT(10) WITH COMPRESSION NOT NULL ," + _
				"	em_y TEXT(10) WITH COMPRESSION NOT NULL ," + _
				"	em_largo TEXT(10) WITH COMPRESSION NOT NULL ," + _
				"	em_alto TEXT(10) WITH COMPRESSION NOT NULL, " + _
				"	html TEXT WITH COMPRESSION NULL ," + _
				"	format TEXT WITH COMPRESSION NULL ," + _
				"	testo TEXT WITH COMPRESSION NULL ," + _
				"	aspcode TEXT WITH COMPRESSION NULL ," + _
				"	RTF TEXT WITH COMPRESSION NULL ," + _
				"	CHECKSUM_STILI TEXT(250) WITH COMPRESSION NULL" + _
				" ) ; "
		case DB_SQL
			Install__FRAMEWORK_CORE__NEXTWEB5 = _
				" CREATE TABLE dbo.tb_tipo ( " + _
				"	id_tip int NOT NULL , " + _
				"	tipo nvarchar (50) NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_css_groups ( " + _
				"	grp_id int IDENTITY (1, 1) NOT NULL , " + _
				"	grp_id_webs int NOT NULL , " + _
				"	grp_name nvarchar (250) NOT NULL , " + _
				"	grp_name_class nvarchar (250) NOT NULL , " + _
				"	grp_for_editor bit NOT NULL , " + _
				"	grp_for_file bit NOT NULL , " + _
				"	grp_checksum nvarchar (250) NOT NULL , " + _
				"	grp_ins_data smalldatetime NOT NULL , " + _
				"	grp_ins_admin_id int NOT NULL , " + _
				"	grp_mod_data smalldatetime NOT NULL , " + _
				"	grp_mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_css_styles ( " + _
				"	style_id int IDENTITY (1, 1) NOT NULL , " + _
				"	style_grp_id int NOT NULL , " + _
				"	style_class nvarchar (20) NOT NULL , " + _
				"	style_pseudoclass nvarchar (20) NULL , " + _
				"	style_font_family nvarchar (250) NULL , " + _
				"	style_font_size float NULL , " + _
				"	style_font_weight nvarchar (20) NULL , " + _
				"	style_font_style nvarchar (20) NULL , " + _
				"	style_color nvarchar (20) NULL , " + _
				"	style_background_color nvarchar (20) NULL , " + _
				"	style_line_height float NULL , " + _
				"	style_letter_spacing float NULL , " + _
				"	style_text_align nvarchar (20) NULL , " + _
				"	style_text_decoration nvarchar (20) NULL , " + _
				"	style_description nvarchar (250) NULL , " + _
				"	style_ins_data smalldatetime NOT NULL , " + _
				"	style_ins_admin_id int NOT NULL , " + _
				"	style_mod_data smalldatetime NOT NULL , " + _
				"	style_mod_admin_id smalldatetime NOT NULL " + _
				"	); " + _
				" CREATE TABLE dbo.tb_objects ( " + _
				"	id_objects int IDENTITY (1, 1) NOT NULL , " + _
				"	id_webs int NOT NULL , " + _
				"	name_objects nvarchar (255) NOT NULL , " + _
				"	identif_objects nvarchar (70) NOT NULL , " + _
				"	param_list ntext NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_pages ( " + _
				"	id_page int IDENTITY (1, 1) NOT NULL , " + _
				"	id_webs int NOT NULL , " + _
				"	id_template int NULL , " + _
				"	lingua varchar (2) NULL , " + _
				"	nomepage nvarchar (250) NOT NULL , " + _
				"	template bit NOT NULL , " + _
				"	SfondoColore nvarchar (7) NULL , " + _
				"	SfondoImmagine nvarchar (255) NULL , " + _
				"	Contatore int NOT NULL , " + _
				"	ContRes smalldatetime NOT NULL , " + _
				"	contUtenti int NOT NULL , " + _
				"	contCrawler int NOT NULL , " + _
				"	contAltro int NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_pagineSito ( " + _
				"	id_pagineSito int IDENTITY (1, 1) NOT NULL , " + _
				"	id_web int NOT NULL , " + _
				"	archiviata bit NOT NULL , " + _
				"	riservata bit NOT NULL , " + _
				"	id_pagDyn_IT int NOT NULL , " + _
				"	id_pagDyn_EN int NULL , " + _
				"	id_pagDyn_FR int NULL , " + _
				"	id_pagDyn_DE int NULL , " + _
				"	id_pagDyn_ES int NULL , " + _
				"	id_pagStage_IT int NOT NULL , " + _
				"	id_pagStage_EN int NULL , " + _
				"	id_pagStage_FR int NULL , " + _
				"	id_pagStage_DE int NULL , " + _
				"	id_pagStage_ES int NULL , " + _
				"	nome_ps_IT nvarchar (255) NOT NULL , " + _
				"	nome_ps_EN nvarchar (255) NULL , " + _
				"	nome_ps_FR nvarchar (255) NULL , " + _
				"	nome_ps_DE nvarchar (255) NULL , " + _
				"	nome_ps_ES nvarchar (255) NULL , " + _
				"	PAGE_keywords_IT ntext NULL , " + _
				"	PAGE_keywords_EN ntext NULL , " + _
				"	PAGE_keywords_FR ntext NULL , " + _
				"	PAGE_keywords_DE ntext NULL , " + _
				"	PAGE_keywords_ES ntext NULL , " + _
				"	PAGE_description_IT ntext NULL , " + _
				"	PAGE_description_EN ntext NULL , " + _
				"	PAGE_description_FR ntext NULL , " + _
				"	PAGE_description_DE ntext NULL , " + _
				"	PAGE_description_ES ntext NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_storico_pages ( " + _
				"	sp_ID int IDENTITY (1, 1) NOT NULL , " + _
				"	sp_sw_id int NOT NULL, " + _
				"	sp_page_id int NOT NULL , " + _
				"	sp_pagineSito_id int NOT NULL , " + _
				"	sp_nomepage nvarchar (250) NOT NULL , " + _
				"	sp_lingua nvarchar (2) NOT NULL , " + _
				"	sp_contatore int NOT NULL , " + _
				"	sp_contUtenti int NOT NULL , " + _
				"	sp_contCrawler int NOT NULL , " + _
				"	sp_contAltro int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_storico_webs ( " + _
				"	sw_ID int IDENTITY (1, 1) NOT NULL , " + _
				"	sw_webs_id int NOT NULL , " + _
				"	sw_data smalldatetime NOT NULL , " + _
				"	sw_contatore int NOT NULL , " + _
				"	sw_contUtenti int NOT NULL , " + _
				"	sw_contCrawler int NOT NULL , " + _
				"	sw_contAltro int NOT NULL , " + _
				"	sw_ins_data smalldatetime NOT NULL , " + _
				"	sw_ins_admin_id smalldatetime NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_webs ( " + _
				"	id_webs int NOT NULL , " + _
				"	nome_webs nvarchar (50) NOT NULL , " + _
				"	lingua_iniziale varchar (2) NOT NULL , " + _
				"	id_home_page int NULL , " + _
				"	id_home_page_riservata int NULL , " + _
				"	sito_in_aggiornamento bit NOT NULL , " + _
				"	sito_in_aggiornamento_pagina int NULL , " + _
				"	sito_in_costruzione bit NOT NULL , " + _
				"	sito_in_costruzione_pagina int NULL , " + _
				"	errore_pagina int NULL , " + _
				"	lingua_EN bit NOT NULL , " + _
				"	lingua_FR bit NOT NULL , " + _
				"	lingua_DE bit NOT NULL , " + _
				"	lingua_ES bit NOT NULL , " + _
				"	titolo_IT nvarchar (255) NOT NULL , " + _
				"	titolo_EN nvarchar (255) NULL , " + _
				"	titolo_FR nvarchar (255) NULL , " + _
				"	titolo_DE nvarchar (255) NULL , " + _
				"	titolo_ES nvarchar (255) NULL , " + _
				"	META_Author ntext NULL , " + _
				"	META_keywords_IT ntext NULL , " + _
				"	META_keywords_EN ntext NULL , " + _
				"	META_keywords_FR ntext NULL , " + _
				"	META_keywords_DE ntext NULL , " + _
				"	META_keywords_ES ntext NULL , " + _
				"	META_description_IT ntext NULL , " + _
				"	META_description_EN ntext NULL , " + _
				"	META_description_FR ntext NULL , " + _
				"	META_description_DE ntext NULL , " + _
				"	META_description_ES ntext NULL , " + _
				"	contatore int NOT NULL , " + _
				"	contRes smalldatetime NOT NULL , " + _
				"	contUtenti int NOT NULL , " + _
				"	contCrawler int NOT NULL , " + _
				"	contAltro int NOT NULL , " + _
				"	ins_data smalldatetime NOT NULL , " + _
				"	ins_admin_id int NOT NULL , " + _
				"	mod_data smalldatetime NOT NULL , " + _
				"	mod_admin_id int NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_layers (" + _
				"	id_lay int IDENTITY (1, 1) NOT NULL ," + _
				"	id_pag int NOT NULL ," + _
				"	id_tipo int NOT NULL ," + _
				"	id_objects int NULL ," + _
				"	tipo_contenuto nvarchar (1) NOT NULL ," + _
				"	z_order int NOT NULL ," + _
				"	nome nvarchar (250) NULL ," + _
				"	visibile bit NOT NULL ," + _
				"	x int NOT NULL ," + _
				"	y int NOT NULL ," + _
				"	largo int NOT NULL ," + _
				"	alto int NOT NULL ," + _
				"	em_x nvarchar (10) NOT NULL ," + _
				"	em_y nvarchar (10) NOT NULL ," + _
				"	em_largo nvarchar (10) NOT NULL ," + _
				"	em_alto nvarchar (10) NOT NULL, " + _
				"	html ntext NULL ," + _
				"	format ntext NULL ," + _
				"	testo ntext NULL ," + _
				"	aspcode ntext NULL ," + _
				"	RTF ntext NULL ," + _
				"	CHECKSUM_STILI nvarchar (250) NULL" + _
				" ); "
	end select
	
	Install__FRAMEWORK_CORE__NEXTWEB5 = Install__FRAMEWORK_CORE__NEXTWEB5 + _
		" ALTER TABLE tb_tipo ADD CONSTRAINT PK_tb_tipo PRIMARY KEY CLUSTERED (id_tip); " + _
		" ALTER TABLE tb_css_groups ADD CONSTRAINT PK_tb_css_groups PRIMARY KEY CLUSTERED (grp_id); " + _
		" ALTER TABLE tb_css_styles ADD CONSTRAINT PK_tb_css_styles PRIMARY KEY CLUSTERED (style_id); " + _
		" ALTER TABLE tb_objects ADD CONSTRAINT PK_tb_objects PRIMARY KEY CLUSTERED (id_objects); " + _
		" ALTER TABLE tb_pages ADD CONSTRAINT PK_tb_pages PRIMARY KEY CLUSTERED (id_page); " + _
		" ALTER TABLE tb_paginesito ADD CONSTRAINT PK_tb_pagineSito PRIMARY KEY CLUSTERED (id_pagineSito); " + _
		" ALTER TABLE tb_storico_pages ADD CONSTRAINT PK_tb_storico_pages PRIMARY KEY CLUSTERED (sp_ID); " + _
		" ALTER TABLE tb_storico_webs ADD CONSTRAINT PK_tb_storico_webs PRIMARY KEY CLUSTERED (sw_ID); " + _
		" ALTER TABLE tb_webs ADD CONSTRAINT PK_tb_webs PRIMARY KEY CLUSTERED (id_webs); " + _
		" ALTER TABLE tb_layers ADD CONSTRAINT PK_tb_layers PRIMARY KEY CLUSTERED (id_lay); " + _
		" ALTER TABLE tb_css_groups ADD CONSTRAINT FK_tb_css_groups_tb_webs " + _
		"	FOREIGN KEY (grp_id_webs) REFERENCES tb_webs (id_webs) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_css_styles ADD CONSTRAINT FK_tb_css_styles_tb_css_groups " + _
		"	FOREIGN KEY (style_grp_id) REFERENCES tb_css_groups (grp_id) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_objects ADD CONSTRAINT FK_tb_objects_tb_webs " + _
		"	FOREIGN KEY (id_webs) REFERENCES tb_webs (id_webs) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_pages ADD " + _
		"	CONSTRAINT FK_tb_pages_tb_webs FOREIGN KEY (id_webs) REFERENCES tb_webs (id_webs) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE, " + _
		"	CONSTRAINT FK_tb_pages_tb_cnt_lingue FOREIGN KEY (lingua) REFERENCES tb_cnt_lingue (lingua_codice) " + _
		"		ON DELETE NO ACTION ON UPDATE NO ACTION ; " + _
		" ALTER TABLE tb_pagineSito ADD CONSTRAINT FK_tb_pagineSito_tb_webs " + _
		"	FOREIGN KEY (id_web) REFERENCES tb_webs (id_webs) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_storico_pages ADD CONSTRAINT FK_tb_storico_pages_tb_storico_webs " + _
		"	FOREIGN KEY (sp_sw_id) REFERENCES tb_storico_webs (sw_ID) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_storico_webs ADD CONSTRAINT FK_tb_storico_webs_tb_webs " + _
		"	FOREIGN KEY (sw_webs_id) REFERENCES tb_webs (id_webs) " + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_webs ADD CONSTRAINT FK_tb_webs_tb_cnt_lingue " + _
		"	FOREIGN KEY (lingua_iniziale) REFERENCES tb_cnt_lingue (lingua_codice) " + _
		"	ON DELETE NO ACTION  ON UPDATE NO ACTION ;" + _
		" ALTER TABLE tb_layers ADD " + _
		"	CONSTRAINT FK_tb_layers_tb_pages FOREIGN KEY (id_pag) REFERENCES tb_pages (id_page) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE , " + _
		"	CONSTRAINT FK_tb_layers_tb_tipo FOREIGN KEY (id_tipo) REFERENCES tb_tipo (id_tip) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" INSERT INTO tb_tipo (id_tip, tipo) VALUES ( 1, 'testo') ; " + _
		" INSERT INTO tb_tipo (id_tip, tipo) VALUES ( 2, 'immagine') ; " + _
		" INSERT INTO tb_tipo (id_tip, tipo) VALUES ( 3, 'animazione') ; " + _
		" INSERT INTO tb_tipo (id_tip, tipo) VALUES ( 4, 'oggetto') ; " + _
		" INSERT INTO tb_tipo (id_tip, tipo) VALUES ( 5, 'testo_strutturato') ; "
	if DB_Type(conn) = DB_SQL then
		Install__FRAMEWORK_CORE__NEXTWEB5 = Install__FRAMEWORK_CORE__NEXTWEB5 + _
			" ALTER TABLE tb_pages ADD " + _
			"	CONSTRAINT FK_tb_pages_tb_pages_template FOREIGN KEY (id_template) REFERENCES tb_pages (id_page) ; " + _
			" ALTER TABLE tb_pagineSito ADD " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_IT FOREIGN KEY (id_pagDyn_IT) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_EN FOREIGN KEY (id_pagDyn_EN) REFERENCES tb_pages (id_page), " + _
			"   CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_FR FOREIGN KEY (id_pagDyn_FR) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_DE FOREIGN KEY (id_pagDyn_DE) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_ES FOREIGN KEY (id_pagDyn_ES) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_Stage_IT FOREIGN KEY (id_pagStage_IT) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_stage_EN FOREIGN KEY (id_pagStage_EN) REFERENCES tb_pages (id_page), " + _
			"   CONSTRAINT FK_tb_pagineSito_tb_pages_stage_FR FOREIGN KEY (id_pagStage_FR) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_stage_DE FOREIGN KEY (id_pagStage_DE) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_pagineSito_tb_pages_stage_ES FOREIGN KEY (id_pagStage_ES) REFERENCES tb_pages (id_page) ; " + _
			" ALTER TABLE tb_pages NOCHECK CONSTRAINT FK_tb_pages_tb_pages_template ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_IT ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_EN ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_FR ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_DE ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_dyn_ES ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_Stage_IT ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_stage_EN ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_stage_FR ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_stage_DE ; " + _
			" ALTER TABLE tb_pagineSito NOCHECK CONSTRAINT FK_tb_pagineSito_tb_pages_stage_ES ; " + _ 
			" ALTER TABLE tb_storico_pages ADD " + _
			"	CONSTRAINT FK_tb_storico_pages_tb_pages FOREIGN KEY (sp_page_id) REFERENCES tb_pages (id_page), " + _
			"	CONSTRAINT FK_tb_storico_pages_tb_pagineSito FOREIGN KEY (sp_pagineSito_id) REFERENCES tb_pagineSito (id_pagineSito) ; " + _
			" ALTER TABLE tb_storico_pages NOCHECK CONSTRAINT FK_tb_storico_pages_tb_pages ; " + _
			" ALTER TABLE tb_storico_pages NOCHECK CONSTRAINT FK_tb_storico_pages_tb_pagineSito ; " + _
			" ALTER TABLE tb_webs ADD " + _
			"	CONSTRAINT FK_tb_webs_tb_pagineSito_pag_home FOREIGN KEY (id_home_page) REFERENCES tb_pagineSito (id_pagineSito) , " + _
			"	CONSTRAINT FK_tb_webs_tb_pagineSito_pag_aggiornamento FOREIGN KEY (sito_in_aggiornamento_pagina) REFERENCES tb_pagineSito (id_pagineSito), " + _
			"	CONSTRAINT FK_tb_webs_tb_pagineSito_pag_costruzione FOREIGN KEY (sito_in_costruzione_pagina) REFERENCES tb_pagineSito (id_pagineSito), " + _
			"	CONSTRAINT FK_tb_webs_tb_pagineSito_pag_errore FOREIGN KEY (errore_pagina) REFERENCES tb_pagineSito (id_pagineSito), " + _
			"	CONSTRAINT FK_tb_webs_tb_pagineSito_pag_riservata FOREIGN KEY (id_home_page_riservata) REFERENCES tb_pagineSito (id_pagineSito) ; " + _
			" ALTER TABLE tb_webs NOCHECK CONSTRAINT FK_tb_webs_tb_pagineSito_pag_home ; " + _
			" ALTER TABLE tb_webs NOCHECK CONSTRAINT FK_tb_webs_tb_pagineSito_pag_aggiornamento ; " + _
			" ALTER TABLE tb_webs NOCHECK CONSTRAINT FK_tb_webs_tb_pagineSito_pag_costruzione ; " + _
			" ALTER TABLE tb_webs NOCHECK CONSTRAINT FK_tb_webs_tb_pagineSito_pag_errore ; " + _
			" ALTER TABLE tb_webs NOCHECK CONSTRAINT FK_tb_webs_tb_pagineSito_pag_riservata ; " + _
			" ALTER TABLE tb_layers ADD CONSTRAINT FK_tb_layers_tb_objects FOREIGN KEY (id_objects) REFERENCES tb_objects (id_objects); " + _
			" ALTER TABLE tb_layers NOCHECK CONSTRAINT FK_tb_layers_tb_objects; "
	end if
end function
'*******************************************************************************************



'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'FUNZIONI PER L'AGGIORNAMENTO DEL FRAMEWORK CORE

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE BASE
'...........................................................................................
'ripulisce directory del framemwork ed aggiorna strutture delle directory
'...........................................................................................
sub Aggiornamento__FRAMEWORK_CORE__pulizia_directory(conn, rs)
	dim fso, FolderUpload, FolderSite, Path, FolderSiteDocs, SubFolder
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	'ripulisce directory temporanee
	CALL ClearTempDir(fso)
	
	'rimuove file inutili dalle directory (qualsiasi directory)
	CALL FileRemove(fso, Application("IMAGE_PATH"), "thumbs.db", true)
	CALL FileRemove(fso, Application("IMAGE_PATH"), "pspbrwse.jbf", true)
	
	'rimuove cartella template (presente solo in alcuni)
	CALL FolderRemove(fso, Application("IMAGE_PATH") & "\template", false)
	
	set FolderUpload = fso.GetFolder(Application("IMAGE_PATH"))
	
	'verifica esistenza delle cartella docs principale
	if not fso.FolderExists(FolderUpload.path + "\docs") then
		CALL fso.CreateFolder(FolderUpload.path + "\docs")
	end if
		
	'scorre tutte le directory dei siti (solo con nome numerico)
	for each FolderSite in FolderUpload.SubFolders
		if isNumeric(FolderSite.name) then
			'rimuove cartella exports
			CALL FolderRemove(fso, FolderSite.path & "\exports", false)
			
			'rimuove cartella temp
			CALL FolderRemove(fso, FolderSite.path & "\temp", false)
			
			'rimuove cartelle vuote da docs interna (eventualmente anche docs)
			CALL Aggiornamento__FRAMEWORK_CORE__pulizia_directory_directory_RemoveEmptyFolders(fso, FolderSite.path + "\docs")
			
			'controllo: se esiste ancora vuol dire che contiene files / cartelle che devono essere spostati.
			if fso.FolderExists(FolderSite.path + "\docs") then
				set FolderSiteDocs = fso.GetFolder(FolderSite.path + "\docs")
				'rinomina eventuali cartele emails da "<id>" a "eml_<id>"
				for each SubFolder in FolderSiteDocs.SubFolders
					if IsNumeric(SubFolder.name) then
						'directory non vuota che contiene i files delle email: deve essere rinominata
						SubFolder.name = "eml_" + SubFolder.name
					elseif instr(1, SubFolder, "pra_", vbTextCompare)>0 then
						'cancella directory residua delle pratiche
						CALL SubFolder.Delete(true)
					end if
				next
				'Copia Directory upload/<az_id>/docs su upload/docs
				CALL FolderSiteDocs.Copy(FolderUpload.path + "\docs", true)
				'rimuove vecchia directory Docs
				CALL FolderSiteDocs.Delete(true)
			end if
		end if
	next
	
	set fso = nothing
end sub


'scorre tutte le directory ed eventualmente cancella tutte le directory vuote, anche quelle figlie.
sub Aggiornamento__FRAMEWORK_CORE__pulizia_directory_directory_RemoveEmptyFolders(fso, BasePath)
	Dim BaseFolder, SubFolders
	if fso.FolderExists(BasePath) then
		set BaseFolder = fso.GetFolder(BasePath)
		if cInteger(BaseFolder.SubFolders.Count)>0 then
			'verifica eventuale rimozione delle cartelle figlie
			for each SubFolder in BaseFolder.SubFolders
				'richiama funzione per tutte le sottodirectory
				CALL Aggiornamento__FRAMEWORK_CORE__pulizia_directory_directory_RemoveEmptyFolders(fso, SubFolder.Path)
			next
		end if
		'ricontrolla la directory per eventuali cancellazioni
		set BaseFolder = fso.GetFolder(BasePath)
		if cInteger(BaseFolder.Files.Count)<=0 AND cInteger(BaseFolder.SubFolders.Count)<=0 then
			'directory vuota: la cancella!
			CALL fso.DeleteFolder(BasePath)
		end if
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 1
'...........................................................................................
'aggiunge campo a categorie delle news e dei link utili per indicazione ordine e pubblicazione
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__1(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__1 = _
				"ALTER TABLE tb_news_categorie ADD " + _
				"	cat_ordine INT NULL, " + _
				"	cat_visibile BIT NULL; " + _
				"ALTER TABLE tb_links_categorie ADD " + _
				"	cat_ordine INT NULL, " + _
				" 	cat_visibile BIT NULL; " + _
				"UPDATE tb_news_categorie SET cat_visibile=1 ; " + _
				"UPDATE tb_links_categorie SET cat_visibile=1 ; " 
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 2
'...........................................................................................
'corregge problemi dichiarazione campi gallery
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__2 = _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_it TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_en TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_fr TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_de TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_es TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_codice TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_it TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_en TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_fr TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_de TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_es TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_descr_it TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_descr_en TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_descr_fr TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_descr_de TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_descr_es TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_it TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_en TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_fr TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_es TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_de TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_it TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_en TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_fr TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_es TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_de TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_Didascalia_it TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_Didascalia_en TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_Didascalia_fr TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_Didascalia_de TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_Didascalia_es TEXT WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_thumb TEXT(250) WITH COMPRESSION NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_zoom TEXT(250) WITH COMPRESSION NULL; "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__2 = _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_it nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_en nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_fr nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_de nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_name_es nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_gallery ALTER COLUMN gallery_codice nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_it nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_en nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_fr nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_de nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_categorieGallery ALTER COLUMN catC_nome_es nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_it nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_en nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_fr nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_es nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_descrittori ALTER COLUMN des_nome_de nvarchar(250) NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_it nvarchar(250) NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_en nvarchar(250) NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_fr nvarchar(250) NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_es nvarchar(250) NULL; " + _
				"ALTER TABLE prel_descrittori_gallery ALTER COLUMN rdi_valore_de nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_thumb nvarchar(250) NULL; " + _
				"ALTER TABLE ptb_Immagini ALTER COLUMN I_zoom nvarchar(250) NULL; "
	end select
	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 3
'...........................................................................................
'uniforma dimensione campo immagine per news
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__3 = _
				"ALTER TABLE tb_news ALTER COLUMN news_img TEXT(250) WITH COMPRESSION NULL ; "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__3 = _
				"ALTER TABLE tb_news ALTER COLUMN news_img nvarchar(250) NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 4
'...........................................................................................
'corregge valori per relazioni NEXT-news o NEXT-link
'per SQL Server corregge anche le relazioni
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__4(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__4 = _
				" UPDATE tb_links SET link_cat_id=NULL WHERE link_cat_id=0 ; " + _
				" UPDATE tb_news SET news_cat_id=NULL WHERE news_cat_id=0 ; "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__4 = _
				" UPDATE tb_links SET link_cat_id=NULL WHERE link_cat_id=0 ; " + _
				" UPDATE tb_news SET news_cat_id=NULL WHERE news_cat_id=0 ; " + _
				" ALTER TABLE tb_links DROP CONSTRAINT FK_tb_links_tb_links_categorie; " + _
				" ALTER TABLE tb_news DROP CONSTRAINT FK_tb_news_tb_news_categorie; " + _
				" ALTER TABLE tb_links ADD CONSTRAINT FK_tb_links_tb_links_categorie " + _
				"	FOREIGN KEY (link_cat_id) REFERENCES tb_links_categorie(cat_id) " + _
				"	ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
				" ALTER TABLE tb_links NOCHECK CONSTRAINT FK_tb_links_tb_links_categorie " + _
				" ALTER TABLE tb_news ADD CONSTRAINT FK_tb_news_tb_news_categorie " + _
				"	FOREIGN KEY (news_cat_id) REFERENCES tb_news_categorie(cat_id) " + _
				"	ON UPDATE NO ACTION ON DELETE NO ACTION ; " + _
				" ALTER TABLE tb_news NOCHECK CONSTRAINT FK_tb_news_tb_news_categorie "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 5
'...........................................................................................
'aggiunge campo per gestione documento allegato a next-news
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__5(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__5 = _
				" ALTER TABLE tb_news ADD COLUMN news_doc TEXT(250) WITH COMPRESSION NULL "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__5 = _
				" ALTER TABLE tb_news ADD news_doc nvarchar(250) NULL "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 6
'...........................................................................................
'corregge problemi con NEXT-gallery per alcuni database access
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__6(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__6 = _
				" ALTER TABLE ptb_gallery ADD COLUMN"+ _
				" 	gallery_name_it2 TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_en2 TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_fr2 TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_es2 TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_de2 TEXT(250) WITH COMPRESSION NULL;"+ _
				" UPDATE ptb_gallery SET"+ _
				"	gallery_name_it2 = gallery_name_it,"+ _
				"	gallery_name_en2 = gallery_name_en,"+ _
				"	gallery_name_fr2 = gallery_name_fr,"+ _
				"	gallery_name_es2 = gallery_name_es,"+ _
				"	gallery_name_de2 = gallery_name_de;"+ _
				" ALTER TABLE ptb_gallery DROP COLUMN"+ _
				"	gallery_name_it,"+ _
				"	gallery_name_en,"+ _
				"	gallery_name_fr,"+ _
				"	gallery_name_es,"+ _
				"	gallery_name_de;"+ _
				" ALTER TABLE ptb_gallery ADD COLUMN"+ _
				" 	gallery_name_it TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_en TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_fr TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_es TEXT(250) WITH COMPRESSION NULL,"+ _
				" 	gallery_name_de TEXT(250) WITH COMPRESSION NULL;"+ _
				" UPDATE ptb_gallery SET"+ _
				"	gallery_name_it = gallery_name_it2,"+ _
				"	gallery_name_en = gallery_name_en2,"+ _
				"	gallery_name_fr = gallery_name_fr2,"+ _
				"	gallery_name_es = gallery_name_es2,"+ _
				"	gallery_name_de = gallery_name_de2;"+ _
				" ALTER TABLE ptb_gallery DROP COLUMN"+ _
				"	gallery_name_it2,"+ _
				"	gallery_name_en2,"+ _
				"	gallery_name_fr2,"+ _
				"	gallery_name_es2,"+ _
				"	gallery_name_de2;"
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__6 = "SELECT * FROM aa_versione"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 7
'...........................................................................................
'NEXT-gallery: collega i descrittori alle categorie
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__7(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__7 = _
				" CREATE TABLE prel_catGallery_descrittori( " + _
				"		rcd_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rcd_categoria_id INT NOT NULL , " + _
				"		rcd_descrittore_id INT NOT NULL , " + _
				"		rcd_ordine INT NULL " + _
				"		); " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT PK_prel_catGallery_descrittori PRIMARY KEY (rcd_id); " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT FK_prel_catGallery_descrittori__ptb_categorieGallery " + _
				"		FOREIGN KEY (rcd_categoria_id) REFERENCES ptb_categorieGallery (catC_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT FK_prel_catGallery_descrittori__ptb_descrittori " + _
				"		FOREIGN KEY (rcd_descrittore_id) REFERENCES ptb_descrittori (des_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" INSERT INTO ptb_categorieGallery(catC_nome_it, catC_foglia, catC_livello, catC_ordine, catC_visibile, catC_albero_visibile)"+ _
				" SELECT TOP 1 'categoria di default', 1, 0, 0, 1, 1 FROM aa_versione"+ _
				" WHERE (SELECT COUNT(*) FROM ptb_categorieGallery) = 0;"+ _
				" UPDATE ptb_categorieGallery SET catC_tipologia_padre_base = catC_id"+ _
				" WHERE ISNULL(catC_tipologia_padre_base);"+ _
				" INSERT INTO prel_catGallery_descrittori(rcd_categoria_id, rcd_descrittore_id, rcd_ordine)"+ _
				" SELECT catC_id, des_id, des_ordine FROM ptb_categorieGallery, ptb_descrittori;"+ _
				" ALTER TABLE ptb_descrittori DROP COLUMN des_ordine;"
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__7 = _
				" CREATE TABLE dbo.prel_catGallery_descrittori( " + _
				"		rcd_id int IDENTITY (1, 1) NOT NULL , " + _
				"		rcd_categoria_id INT NOT NULL , " + _
				"		rcd_descrittore_id INT NOT NULL , " + _
				"		rcd_ordine INT NULL " + _
				"		); " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT PK_prel_catGallery_descrittori PRIMARY KEY (rcd_id); " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT FK_prel_catGallery_descrittori__ptb_categorieGallery " + _
				"		FOREIGN KEY (rcd_categoria_id) REFERENCES ptb_categorieGallery (catC_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" ALTER TABLE prel_catGallery_descrittori ADD CONSTRAINT FK_prel_catGallery_descrittori__ptb_descrittori " + _
				"		FOREIGN KEY (rcd_descrittore_id) REFERENCES ptb_descrittori (des_id) " + _
				"		ON UPDATE CASCADE ON DELETE CASCADE; " + _
				" INSERT INTO ptb_categorieGallery(catC_nome_it, catC_foglia, catC_livello, catC_ordine, catC_visibile, catC_albero_visibile)"+ _
				" SELECT TOP 1 'categoria di default', 1, 0, 0, 1, 1 FROM aa_versione"+ _
				" WHERE (SELECT COUNT(*) FROM ptb_categorieGallery) = 0;"+ _
				" UPDATE ptb_categorieGallery SET catC_tipologia_padre_base = catC_id"+ _
				" WHERE catC_tipologia_padre_base IS NULL;"+ _
				" INSERT INTO prel_catGallery_descrittori(rcd_categoria_id, rcd_descrittore_id, rcd_ordine)"+ _
				" SELECT catC_id, des_id, des_ordine FROM ptb_categorieGallery, ptb_descrittori;"+ _
				" ALTER TABLE ptb_descrittori DROP COLUMN des_ordine;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 8
'...........................................................................................
'NEXT-gallery: conclude il collegamento tra i descrittori alle categorie
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__8(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__8 = _
				" UPDATE ptb_gallery SET gallery_idcategoria = 1"+ _
				" WHERE ISNULL(gallery_idcategoria)"
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__8 = _
				" UPDATE ptb_gallery SET gallery_idcategoria = 1"+ _
				" WHERE gallery_idcategoria IS NULL"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 9
'...........................................................................................
'NEXT-passport: aggiunge campi per gestione permessi esterni
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__9(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__9 = _
				" ALTER TABLE tb_siti ADD"+ _
				" 		sito_prmEsterni_admin TEXT(250) WITH COMPRESSION NULL,"+ _
				"		sito_prmEsterni_sito TEXT(250) WITH COMPRESSION NULL"
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__9 = _
				" ALTER TABLE tb_siti ADD"+ _
				" 		sito_prmEsterni_admin NVARCHAR(250) NULL,"+ _
				"		sito_prmEsterni_sito NVARCHAR(250) NULL"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 10
'...........................................................................................
'NEXT-com: aggiunge campi di traduzione tipi di recapiti
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__10(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__10 = _
				" ALTER TABLE tb_TipNumeri ADD "+ _
				" 		nome_tiponumero_it TEXT(250) WITH COMPRESSION NULL, " + _
				" 		nome_tiponumero_en TEXT(250) WITH COMPRESSION NULL, " + _
				" 		nome_tiponumero_fr TEXT(250) WITH COMPRESSION NULL, " + _
				" 		nome_tiponumero_de TEXT(250) WITH COMPRESSION NULL, " + _
				" 		nome_tiponumero_es TEXT(250) WITH COMPRESSION NULL; " + _
				" ALTER TABLE tb_tipnumeri DROP tiponumero; " 
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__10 = _
				" ALTER TABLE tb_TipNumeri ADD "+ _
				" 		nome_tiponumero_it NVARCHAR(250) NULL, " + _
				" 		nome_tiponumero_en NVARCHAR(250) NULL, " + _
				" 		nome_tiponumero_fr NVARCHAR(250) NULL, " + _
				" 		nome_tiponumero_de NVARCHAR(250) NULL, " +  _
				" 		nome_tiponumero_es NVARCHAR(250) NULL; " + _
				" ALTER TABLE tb_tipnumeri DROP COLUMN tiponumero; " 
	end select
	Aggiornamento__FRAMEWORK_CORE__10 = Aggiornamento__FRAMEWORK_CORE__10 + _
			    " UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Telefono', " + _
				"						  nome_tiponumero_en = 'Telephone', " + _
				"						  nome_tiponumero_fr = 'T&eacute;;l&eacute;;phone', " + _
				"						  nome_tiponumero_de = 'Telefon', " + _
				"						  nome_tiponumero_es = 'Tel&eacute;;fono' " + _
				" WHERE id_TipoNumero=1  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Telefono ufficio', " + _
				"						  nome_tiponumero_en = 'Office telephone', " + _
				"						  nome_tiponumero_fr = 'T&eacute;;l&eacute;;phone d''Office', " + _
				"						  nome_tiponumero_de = 'B&uuml;;rotelefon', " + _
				"						  nome_tiponumero_es = 'Office tel&eacute;;fono' " + _
				" WHERE id_TipoNumero=2  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Telefono cellulare', " + _
				"						  nome_tiponumero_en = 'Mobile phone', " + _
				"						  nome_tiponumero_fr = 'T&eacute;;l&eacute;;phone cellulaire', " + _
				"						  nome_tiponumero_de = 'Zellular Telefon', " + _
				"						  nome_tiponumero_es = 'Tel&eacute;;fono celular' " + _
				" WHERE id_TipoNumero=3  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Telefono casa', " + _
				"						  nome_tiponumero_en = 'Home telephone', " + _
				"						  nome_tiponumero_fr = 'T&eacute;;l&eacute;;phone de la maison', " + _
				"						  nome_tiponumero_de = 'Haupttelefon', " + _
				"						  nome_tiponumero_es = 'Tel&eacute;;fono casero' " + _
				" WHERE id_TipoNumero=4  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Fax', " + _
				"						  nome_tiponumero_en = 'Fax', " + _
				"						  nome_tiponumero_fr = 'Fax', " + _
				"						  nome_tiponumero_de = 'Fax', " + _
				"						  nome_tiponumero_es = 'Fax' " + _
				" WHERE id_TipoNumero=5  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Email', " + _
				"						  nome_tiponumero_en = 'Email', " + _
				"						  nome_tiponumero_fr = 'Email', " + _
				"						  nome_tiponumero_de = 'Email', " + _
				"						  nome_tiponumero_es = 'Email' " + _
				" WHERE id_TipoNumero=6  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Sito internet', " + _
				"						  nome_tiponumero_en = 'Web site', " + _
				"						  nome_tiponumero_fr = 'Site Web', " + _
				"						  nome_tiponumero_de = 'Web site', " + _
				"						  nome_tiponumero_es = 'Web site' " + _
				" WHERE id_TipoNumero=7  ; " + _
			    " UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Telefono', " + _
				"						  nome_tiponumero_en = 'Telephone', " + _
				"						  nome_tiponumero_fr = 'T&eacute;;l&eacute;;phone ', " + _
				"						  nome_tiponumero_de = 'Telefon', " + _
				"						  nome_tiponumero_es = 'Tel&eacute;;fono' " + _
				" WHERE id_TipoNumero=8  ; " + _
				" UPDATE tb_tipNumeri SET nome_tiponumero_it = 'Fax', " + _
				"						  nome_tiponumero_en = 'Fax', " + _
				"						  nome_tiponumero_fr = 'Fax', " + _
				"						  nome_tiponumero_de = 'Fax', " + _
				"						  nome_tiponumero_es = 'Fax' " + _
				" WHERE id_TipoNumero=9  ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 11
'...........................................................................................
'NEXT-com: aggiunge campo per gestione flag di protezione privacy per non pubblicazione
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__11(conn)
	Aggiornamento__FRAMEWORK_CORE__11 = _
		" ALTER TABLE tb_ValoriNumeri ADD " + _
		" 		protetto_privacy BIT NULL; " + _
		" UPDATE tb_valorinumeri SET protetto_privacy=0 ; " + _
		" ALTER TABLE tb_ValoriNumeri ALTER COLUMN protetto_privacy BIT NOT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 12
'...........................................................................................
'corregge valore COLLATE su tutte le colonne testuali non FOREIGN KEY o PRIMARY KEY
'...........................................................................................
sub AggiornamentoSpeciale__FRAMEWORK_CORE__12(DB, rs, versione)
	
	CALL DB.execute("SELECT * FROM AA_versione", versione)
	
	DB.SqlServer_COLLATE_REBUILD(versione)
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 13
'...........................................................................................
'NEXT-com: aggiunge campo password al contatto per sicurezza in contattaci.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__13(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__13 = _
				" ALTER TABLE tb_indirizzario ADD" + _
				" 		codiceInserimento TEXT(50) NULL;"
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__13 = _
				" ALTER TABLE tb_indirizzario ADD" + _
				" 		codiceInserimento NVARCHAR(50) NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 14
'...........................................................................................
'NEXT-web5: rimuove campi per tracciatura modifiche mettendo quelli standard
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__14(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__14 = _
				" ALTER TABLE tb_css_groups DROP COLUMN grp_ins_data ;" + _
				" ALTER TABLE tb_css_groups DROP COLUMN grp_ins_admin_id ;" + _
				" ALTER TABLE tb_css_groups DROP COLUMN grp_mod_data ;" + _
				" ALTER TABLE tb_css_groups DROP COLUMN grp_mod_admin_id ;" + _
				" ALTER TABLE tb_css_styles DROP COLUMN style_ins_data ;" + _
				" ALTER TABLE tb_css_styles DROP COLUMN style_ins_admin_id ;" + _
				" ALTER TABLE tb_css_styles DROP COLUMN style_mod_data ;" + _
				" ALTER TABLE tb_css_styles DROP COLUMN style_mod_admin_id ;" + _
				" ALTER TABLE tb_objects DROP COLUMN ins_data ;" + _
				" ALTER TABLE tb_objects DROP COLUMN ins_admin_id ;" + _
				" ALTER TABLE tb_objects DROP COLUMN mod_data ;" + _
				" ALTER TABLE tb_objects DROP COLUMN mod_admin_id ;" + _
				" ALTER TABLE tb_pages DROP COLUMN ins_data ;" + _
				" ALTER TABLE tb_pages DROP COLUMN ins_admin_id ;" + _
				" ALTER TABLE tb_pages DROP COLUMN mod_data ;" + _
				" ALTER TABLE tb_pages DROP COLUMN mod_admin_id ;" + _
				" ALTER TABLE tb_pagineSito DROP COLUMN ins_data ;" + _
				" ALTER TABLE tb_pagineSito DROP COLUMN ins_admin_id ;" + _
				" ALTER TABLE tb_pagineSito DROP COLUMN mod_data ;" + _
				" ALTER TABLE tb_pagineSito DROP COLUMN mod_admin_id ;" + _
				" ALTER TABLE tb_webs DROP COLUMN ins_data ;" + _
				" ALTER TABLE tb_webs DROP COLUMN ins_admin_id ;" + _
				" ALTER TABLE tb_webs DROP COLUMN mod_data ;" + _
				" ALTER TABLE tb_webs DROP COLUMN mod_admin_id ;" + _
				" ALTER TABLE tb_storico_webs DROP COLUMN sw_ins_data ; " + _
				" ALTER TABLE tb_storico_webs DROP COLUMN sw_ins_admin_id ; " + _
				" ALTER TABLE tb_css_groups ADD " + _
				"	grp_insData smalldatetime NULL , " + _
				"	grp_insAdmin_id int NULL , " + _
				"	grp_modData smalldatetime NULL , " + _
				"	grp_modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_css_styles ADD " + _
				"	style_insData smalldatetime NULL , " + _
				"	style_insAdmin_id int NULL , " + _
				"	style_modData smalldatetime NULL , " + _
				"	style_modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_objects ADD " + _
				"	insData smalldatetime NULL , " + _
				"	insAdmin_id int NULL , " + _
				"	modData smalldatetime NULL , " + _
				"	modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_pages ADD " + _
				"	insData smalldatetime NULL , " + _
				"	insAdmin_id int NULL , " + _
				"	modData smalldatetime NULL , " + _
				"	modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_pagineSito ADD " + _
				"	insData smalldatetime NULL , " + _
				"	insAdmin_id int NULL , " + _
				"	modData smalldatetime NULL , " + _
				"	modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_webs ADD " + _
				"	insData smalldatetime NULL , " + _
				"	insAdmin_id int NULL , " + _
				"	modData smalldatetime NULL , " + _
				"	modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_storico_webs ADD " + _
				"	sw_insData smalldatetime NULL , " + _
				"	sw_insAdmin_id int NULL, " + _
				"	sw_modData smalldatetime NULL , " + _
				"	sw_modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_css_groups ALTER COLUMN grp_insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_css_groups ALTER COLUMN grp_insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_css_groups ALTER COLUMN grp_modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_css_groups ALTER COLUMN grp_modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_css_styles ALTER COLUMN style_insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_css_styles ALTER COLUMN style_insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_css_styles ALTER COLUMN style_modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_css_styles ALTER COLUMN style_modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_objects ALTER COLUMN insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_objects ALTER COLUMN insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_objects ALTER COLUMN modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_objects ALTER COLUMN modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_pages ALTER COLUMN insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_pages ALTER COLUMN insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_pages ALTER COLUMN modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_pages ALTER COLUMN modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_pagineSito ALTER COLUMN insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_pagineSito ALTER COLUMN insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_pagineSito ALTER COLUMN modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_pagineSito ALTER COLUMN modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_webs ALTER COLUMN insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_webs ALTER COLUMN insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_webs ALTER COLUMN modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_webs ALTER COLUMN modAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_storico_webs ALTER COLUMN sw_insData smalldatetime NOT NULL; " + _
				" ALTER TABLE tb_storico_webs ALTER COLUMN sw_insAdmin_id int NOT NULL; " + _
				" ALTER TABLE tb_storico_webs ALTER COLUMN sw_modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_storico_webs ALTER COLUMN sw_modAdmin_id int NOT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 15
'...........................................................................................
'NEXT-web 5.0: aggiunge gestione indice generale e tagging
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__15(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__15 = _
				" CREATE TABLE tb_index ( " + _
				"	idx_id int IDENTITY (1, 1) NOT NULL , " + _
				"	idx_F_table_id INT NOT NULL, " + _
				"	idx_F_key_id INT NOT NULL, " + _
				"	idx_titolo_IT TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	idx_titolo_EN TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_titolo_FR TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_titolo_DE TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_titolo_ES TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_link_url TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	idx_link_tipo TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	idx_foglia BIT NOT NULL, " + _
				"	idx_livello INT NOT NULL, " + _
				"	idx_ordine INT NULL, " + _
				"	idx_ordine_assoluto TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_padre_id INT NULL, " + _
				"	idx_tipologia_padre_base INT NULL, " + _
				"	idx_tipologie_padre_lista TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_foto_thumb TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_foto_zoom TEXT(255) WITH COMPRESSION NULL, " + _
				"	idx_chiave_IT TEXT(100) WITH COMPRESSION NULL, " + _
				"	idx_chiave_EN TEXT(100) WITH COMPRESSION NULL, " + _
				"	idx_chiave_FR TEXT(100) WITH COMPRESSION NULL, " + _
				"	idx_chiave_DE TEXT(100) WITH COMPRESSION NULL, " + _
				"	idx_chiave_ES TEXT(100) WITH COMPRESSION NULL, " + _
				"	idx_descrizione_IT TEXT WITH COMPRESSION NULL, " + _
				"	idx_descrizione_EN TEXT WITH COMPRESSION NULL, " + _
				"	idx_descrizione_FR TEXT WITH COMPRESSION NULL, " + _
				"	idx_descrizione_DE TEXT WITH COMPRESSION NULL, " + _
				"	idx_descrizione_ES TEXT WITH COMPRESSION NULL " + _
				" ); " + _
				" CREATE TABLE tb_siti_tabelle ( " + _
				"	tab_id INT IDENTITY(1,1) NOT NULL, " + _
				"	tab_sito_id INT NOT NULL, " + _
				"	tab_titolo TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_name TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_field_chiave TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_field_titolo TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_field_descrizione TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_field_foto_thumb TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_field_foto_zoom TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tab_Url_Scheda_Admin TEXT(255) WITH COMPRESSION NOT NULL " + _
				" ); " + _
				" CREATE TABLE tb_index_tags ( " + _
				"	tag_id INT IDENTITY(1,1) NOT NULL, " + _
				"	tag_IT TEXT(255) WITH COMPRESSION NOT NULL, " + _
				"	tag_EN TEXT(255) WITH COMPRESSION NULL, " + _
				"	tag_FR TEXT(255) WITH COMPRESSION NULL, " + _
				"	tag_DE TEXT(255) WITH COMPRESSION NULL, " + _
				"	tag_ES TEXT(255) WITH COMPRESSION NULL " + _
				" ); "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__15 = _
				" CREATE TABLE dbo.tb_index ( " + _
				"	idx_id int IDENTITY (1, 1) NOT NULL , " + _
				"	idx_F_table_id INT NOT NULL, " + _
				"	idx_F_key_id INT NOT NULL, " + _
				"	idx_titolo_IT nvarchar(255) NOT NULL, " + _
				"	idx_titolo_EN nvarchar(255) NULL, " + _
				"	idx_titolo_FR nvarchar(255) NULL, " + _
				"	idx_titolo_DE nvarchar(255) NULL, " + _
				"	idx_titolo_ES nvarchar(255) NULL, " + _
				"	idx_link_url nvarchar(255) NULL, " + _
				"	idx_link_tipo nvarchar(255) NULL, " + _
				"	idx_foglia BIT NOT NULL, " + _
				"	idx_livello INT NOT NULL, " + _
				"	idx_ordine INT NULL, " + _
				"	idx_ordine_assoluto nvarchar(255) NULL, " + _
				"	idx_padre_id INT NULL, " + _
				"	idx_tipologia_padre_base INT NULL, " + _
				"	idx_tipologie_padre_lista nvarchar(255) NULL, " + _
				"	idx_foto_thumb nvarchar(255) NULL, " + _
				"	idx_foto_zoom nvarchar(255) NULL, " + _
				"	idx_chiave_IT nvarchar(100) NULL, " + _
				"	idx_chiave_EN nvarchar(100) NULL, " + _
				"	idx_chiave_FR nvarchar(100) NULL, " + _
				"	idx_chiave_DE nvarchar(100) NULL, " + _
				"	idx_chiave_ES nvarchar(100) NULL, " + _
				"	idx_descrizione_IT ntext NULL, " + _
				"	idx_descrizione_EN ntext NULL, " + _
				"	idx_descrizione_FR ntext NULL, " + _
				"	idx_descrizione_DE ntext NULL, " + _
				"	idx_descrizione_ES ntext NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_siti_tabelle ( " + _
				"	tab_id INT IDENTITY(1,1) NOT NULL, " + _
				"	tab_sito_id INT NOT NULL, " + _
				"	tab_titolo nvarchar(255) NOT NULL, " + _
				"	tab_name nvarchar(255) NOT NULL, " + _
				"	tab_field_chiave nvarchar(255) NOT NULL, " + _
				"	tab_field_titolo nvarchar(255) NOT NULL, " + _
				"	tab_field_descrizione nvarchar(255) NOT NULL, " + _
				"	tab_field_foto_thumb nvarchar(255) NOT NULL, " + _
				"	tab_field_foto_zoom nvarchar(255) NOT NULL, " + _
				"	tab_Url_Scheda_Admin nvarchar(255) NOT NULL " + _
				" ); " + _
				" CREATE TABLE dbo.tb_index_tags ( " + _
				"	tag_id INT IDENTITY(1,1) NOT NULL, " + _
				"	tag_IT nvarchar(255) NOT NULL, " + _
				"	tag_EN nvarchar(255) NULL, " + _
				"	tag_FR nvarchar(255) NULL, " + _
				"	tag_DE nvarchar(255) NULL, " + _
				"	tag_ES nvarchar(255) NULL " + _
				" ); "
	end select
	Aggiornamento__FRAMEWORK_CORE__15 = Aggiornamento__FRAMEWORK_CORE__15 + _
				" CREATE TABLE " & SQL_Dbo(Conn) & "rel_index ( " + _
				"	rii_id INT IDENTITY(1,1) NOT NULL, " + _
				"	rii_index_id INT NOT NULL, " + _
				"	rii_collegato_id INT NOT NULL " + _
				" ); " + _
				" CREATE TABLE " & SQL_Dbo(Conn) & "rel_index_tags ( " + _
				"	rit_id INT IDENTITY(1,1) NOT NULL, " + _
				"	rit_index_id INT NOT NULL, " + _
				"	rit_tag_id INT NOT NULL " + _
				" ); " + _
				" ALTER TABLE tb_index ADD CONSTRAINT PK_tb_index PRIMARY KEY CLUSTERED (idx_id); " + _
				" ALTER TABLE tb_index_tags ADD CONSTRAINT PK_tb_index_tags PRIMARY KEY CLUSTERED (tag_id); " + _
				" ALTER TABLE tb_siti_tabelle ADD CONSTRAINT PK_tb_siti_tabelle PRIMARY KEY CLUSTERED (tab_id); " + _
				" ALTER TABLE rel_index ADD CONSTRAINT PK_rel_index PRIMARY KEY CLUSTERED (rii_id); " + _
				" ALTER TABLE rel_index_tags ADD CONSTRAINT PK_rel_index_tags PRIMARY KEY CLUSTERED (rit_id); " + _
				" ALTER TABLE tb_siti_tabelle ADD CONSTRAINT FK_tb_siti_tabelle_tb_siti " + _
				"	FOREIGN KEY (tab_sito_id) REFERENCES tb_siti (id_sito) " + _
				"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
				" ALTER TABLE tb_index ADD CONSTRAINT FK_tb_index_tb_siti_tabelle " + _
				"	FOREIGN KEY (idx_F_table_id) REFERENCES tb_siti_tabelle (tab_id) " + _
				"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
				" ALTER TABLE rel_index ADD " + _
				"	CONSTRAINT FK_rel_index_tb_index__index FOREIGN KEY (rii_index_id) REFERENCES tb_index(idx_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " +_
				"	CONSTRAINT FK_rel_index_tb_index__collegato FOREIGN KEY (rii_collegato_id) REFERENCES tb_index(idx_id) ; " +_
				" ALTER TABLE rel_index_tags ADD " + _
				"	CONSTRAINT FK_rel_index_tags_tb_index FOREIGN KEY (rit_index_id) REFERENCES tb_index(idx_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE, " +_
				"	CONSTRAINT FK_rel_index_tags_tb_index_tags FOREIGN KEY (rit_tag_id) REFERENCES tb_index_tags(tag_id) " + _
				"		ON DELETE CASCADE ON UPDATE CASCADE ; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__15 = Aggiornamento__FRAMEWORK_CORE__15 + _
				" ALTER TABLE tb_index ADD " + _
				"	CONSTRAINT FK_tb_index_tb_index_padre FOREIGN KEY (idx_padre_id) REFERENCES tb_index (idx_id), " + _
				"	CONSTRAINT FK_tb_index_tb_index_padre_base FOREIGN KEY (idx_tipologia_padre_base) REFERENCES tb_index(idx_id) ; " +_
				" ALTER TABLE tb_index NOCHECK CONSTRAINT FK_tb_index_tb_index_padre; " + _
				" ALTER TABLE tb_index NOCHECK CONSTRAINT FK_tb_index_tb_index_padre_base; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 16
'...........................................................................................
'NEXT-web 5.0: corregge valori not nul su inserimento tb_paginesito
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__16(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__16 = _
				" ALTER TABLE tb_paginesito ALTER COLUMN id_pagDyn_IT INT NULL; " + _
				" ALTER TABLE tb_paginesito ALTER COLUMN id_pagStage_IT INT NULL; "
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 17
'...........................................................................................
'NEXT-web 5.0: aggiunge struttura di gestione permessi del next-index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__17(conn)
	Aggiornamento__FRAMEWORK_CORE__17 = _
				" CREATE TABLE " & SQL_Dbo(Conn) & "rel_index_admin( " + _
				"	ria_id INT IDENTITY(1,1) NOT NULL, " + _
				"	ria_index_id INT NOT NULL, " + _
				"	ria_admin_id INT NOT NULL, " + _
				"	ria_modifica BIT NOT NULL, " + _
				"	ria_pubblica BIT NOT NULL " + _
				" ) ; " + _
				" ALTER TABLE rel_index_admin ADD CONSTRAINT FK_rel_index_admin_tb_index " + _
				"	FOREIGN KEY (ria_index_id) REFERENCES tb_index (idx_id) " + _
				"	ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
				" ALTER TABLE rel_index_admin ADD CONSTRAINT FK_rel_index_admin_tb_admin " + _
				"	FOREIGN KEY (ria_admin_id) REFERENCES tb_admin (id_admin) " + _
				"	ON DELETE CASCADE  ON UPDATE CASCADE ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 18
'...........................................................................................
'NEXT-web 5.0: aggiunge campi per indicizzazione scadenza e visibilita' articolo
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__18(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__18 = _
					" ALTER TABLE tb_index ADD " + _
					"	idx_visibile BIT NOT NULL, " + _
					"	idx_data_pubblicazione SMALLDATETIME NULL, " + _
					"	idx_data_scadenza SMALLDATETIME NULL ; " + _
					" ALTER TABLE tb_siti_tabelle ADD " + _
					"	tab_field_visibile TEXT(255) WITH COMPRESSION NULL, " + _
					"	tab_field_data_pubblicazione TEXT(255) WITH COMPRESSION NULL, " + _
					"	tab_field_data_scadenza TEXT(255) WITH COMPRESSION NULL ; "
		case DB_SQL
			Aggiornamento__FRAMEWORK_CORE__18 = _
					" ALTER TABLE tb_index ADD " + _
					"	idx_visibile BIT NULL, " + _
					"	idx_data_pubblicazione SMALLDATETIME NULL, " + _
					"	idx_data_scadenza SMALLDATETIME NULL ; " + _
					" ALTER TABLE tb_index ALTER COLUMN idx_visibile BIT NOT NULL ; " + _
					" ALTER TABLE tb_siti_tabelle ADD " + _
					"	tab_field_visibile nvarchar(255) NULL, " + _
					"	tab_field_data_pubblicazione nvarchar(255) NULL, " + _
					"	tab_field_data_scadenza nvarchar(255) NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 19
'...........................................................................................
'NEXT-web 5.0: 	rimuove flag di gestioni permessi dell'index
'				aggiunge campi per tracciatura modifiche dell'index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__19(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__19 = _
				" ALTER TABLE rel_index_admin DROP COLUMN ria_modifica ;" + _
				" ALTER TABLE rel_index_admin DROP COLUMN ria_pubblica ;" + _
				" ALTER TABLE tb_index ADD " + _
				"	idx_insData smalldatetime NULL , " + _
				"	idx_insAdmin_id int NULL, " + _
				"	idx_modData smalldatetime NULL , " + _
				"	idx_modAdmin_id int NULL ; " + _
				" ALTER TABLE tb_index ALTER COLUMN idx_insData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_index ALTER COLUMN idx_insAdmin_id int NOT NULL ; " + _
				" ALTER TABLE tb_index ALTER COLUMN idx_modData smalldatetime NOT NULL ; " + _
				" ALTER TABLE tb_index ALTER COLUMN idx_modAdmin_id int NOT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 20
'...........................................................................................
'NEXT-web 5.0: 	aggiunge flag di indicazione del tipo di contenuto all'index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__20(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__20 = _
				" ALTER TABLE tb_index ADD " + _
				"	idx_contenuto BIT NULL; " + _
				" ALTER TABLE tb_index ALTER COLUMN idx_contenuto BIT NOT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 21
'...........................................................................................
'NEXT-web 5.0: rimuove gestione indice e contenuti fatta finora e la ricrea nuova versione.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__21(conn)
	
	'rimozione relazioni ed oggetti
	
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__21 =  + _
			" ALTER TABLE tb_index DROP CONSTRAINT FK_tb_index_tb_index_padre; " + _
			" ALTER TABLE tb_index DROP CONSTRAINT FK_tb_index_tb_index_padre_base; "
	end if
	Aggiornamento__FRAMEWORK_CORE__21 = Aggiornamento__FRAMEWORK_CORE__21 + _
		" ALTER TABLE tb_siti_tabelle DROP CONSTRAINT FK_tb_siti_tabelle_tb_siti ; " + _
		" ALTER TABLE tb_index DROP CONSTRAINT FK_tb_index_tb_siti_tabelle ; " + _
		" ALTER TABLE rel_index DROP CONSTRAINT FK_rel_index_tb_index__index ;" +_
		" ALTER TABLE rel_index DROP CONSTRAINT FK_rel_index_tb_index__collegato ; " +_
		" ALTER TABLE rel_index_tags DROP CONSTRAINT FK_rel_index_tags_tb_index ; " +_
		" ALTER TABLE rel_index_tags DROP CONSTRAINT FK_rel_index_tags_tb_index_tags; " + _
		" ALTER TABLE rel_index_admin DROP CONSTRAINT FK_rel_index_admin_tb_index ; " + _
		" ALTER TABLE rel_index_admin DROP CONSTRAINT FK_rel_index_admin_tb_admin ; " + _
		" ALTER TABLE tb_index DROP CONSTRAINT PK_tb_index; " + _
		" ALTER TABLE tb_index_tags DROP CONSTRAINT PK_tb_index_tags; " + _
		" ALTER TABLE tb_siti_tabelle DROP CONSTRAINT PK_tb_siti_tabelle; " + _
		" ALTER TABLE rel_index DROP CONSTRAINT PK_rel_index; " + _
		" ALTER TABLE rel_index_tags DROP CONSTRAINT PK_rel_index_tags; " + _
		DropObject(conn, "tb_index", "TABLE") + _
		DropObject(conn, "rel_index", "TABLE") + _
		DropObject(conn, "tb_index_tags", "TABLE") + _
		DropObject(conn, "tb_siti_tabelle", "TABLE") + _
		DropObject(conn, "rel_index_tags", "TABLE") + _
		DropObject(conn, "rel_index_admin", "TABLE")
		
	'aggiunge nuova versione indice
	Aggiornamento__FRAMEWORK_CORE__21 = Aggiornamento__FRAMEWORK_CORE__21 + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "rel_contents ( " + _
		"	rcc_id int IDENTITY (1, 1) NOT NULL , " + _
		"	rcc_content_id int NOT NULL , " + _
		"	rcc_correlato_id int NOT NULL , " + _
		"	CONSTRAINT PK_rel_contents PRIMARY KEY CLUSTERED (rcc_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "rel_contents_tags ( " + _
		"	rct_id int IDENTITY (1, 1) NOT NULL , " + _
		"	rct_content_id int NOT NULL , " + _
		"	rct_tag_id int NOT NULL , " + _
		"	CONSTRAINT PK_rel_contents_tags PRIMARY KEY CLUSTERED (rct_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "rel_index_admin ( " + _
		"	ria_id int IDENTITY (1, 1) NOT NULL , " + _
		"	ria_index_id int NOT NULL , " + _
		"	ria_admin_id int NOT NULL , " + _
		"	CONSTRAINT PK_rel_index_admin PRIMARY KEY CLUSTERED (ria_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_contents ( " + _
		"	co_id int IDENTITY (1, 1) NOT NULL , " + _
		"	co_F_table_id int NOT NULL , " + _
		"	co_F_key_id int NOT NULL , " + _
		"	co_titolo_IT " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	co_titolo_EN " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_titolo_FR " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_titolo_DE " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_titolo_ES " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_ordine int NULL , " + _
		"	co_foto_thumb " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_foto_zoom " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_chiave_IT " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_chiave_EN " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_chiave_FR " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_chiave_DE " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_chiave_ES " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_descrizione_IT " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	co_descrizione_EN " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	co_descrizione_FR " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	co_descrizione_DE " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	co_descrizione_ES " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	co_visibile bit NOT NULL , " + _
		"	co_data_pubblicazione smalldatetime NULL , " + _
		"	co_data_scadenza smalldatetime NULL , " + _
		"	co_insData smalldatetime NOT NULL , " + _
		"	co_insAdmin_id int NOT NULL , " + _
		"	co_modData smalldatetime NOT NULL , " + _
		"	co_modAdmin_id int NOT NULL , " + _
		"	CONSTRAINT PK_tb_contents PRIMARY KEY  CLUSTERED (co_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_contents_tags ( " + _
		"	tag_id int IDENTITY (1, 1) NOT NULL , " + _
		"	tag_IT " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tag_EN " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tag_FR " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tag_DE " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tag_ES " + SQL_CharField(Conn, 255) + " NULL , " + _ 
		" 	CONSTRAINT PK_tb_contents_tags PRIMARY KEY  CLUSTERED (tag_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_contents_index ( " + _
		"	idx_id int IDENTITY (1, 1) NOT NULL , " + _
		"	idx_content_id int NOT NULL , " + _
		"	idx_link_url_IT " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_link_url_EN " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_link_url_FR " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_link_url_DE " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_link_url_ES " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_link_tipo " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_foglia bit NOT NULL , " + _
		"	idx_livello int NOT NULL , " + _
		"	idx_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_visibile_assoluto bit NULL , " + _
		"	idx_padre_id int NULL , " + _
		"	idx_tipologia_padre_base int NULL , " + _
		"	idx_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	idx_insData smalldatetime NOT NULL , " + _
		"	idx_insAdmin_id int NOT NULL , " + _
		"	idx_modData smalldatetime NOT NULL , " + _
		"	idx_modAdmin_id int NOT NULL , " + _
		"	CONSTRAINT PK_tb_contents_index PRIMARY KEY  CLUSTERED (idx_id) " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_siti_tabelle ( " + _
		"	tab_id int IDENTITY (1, 1) NOT NULL , " + _
		"	tab_sito_id int NOT NULL , " + _
		"	tab_titolo " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_name " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_chiave " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_titolo " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_descrizione " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_foto_thumb " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_foto_zoom " + SQL_CharField(Conn, 255) + " NOT NULL , " + _
		"	tab_field_visibile " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tab_field_ordine " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tab_field_data_pubblicazione " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tab_field_data_scadenza " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	CONSTRAINT PK_tb_siti_tabelle PRIMARY KEY  CLUSTERED (tab_id) " + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_siti_tabelle_pubblicazioni ( " + _
		"	pub_id int IDENTITY (1, 1) NOT NULL , " + _
		"	pub_titolo " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	pub_tabella_id int NOT NULL , " + _
		"	pub_pagina_id int NULL , " + _
		"	pub_parametro " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	pub_index_padre_base_id int NOT NULL , " + _
		" 	pub_filtro_pubblicazione " + SQL_CharField(Conn, 0) + " NULL , " + _
		"	pub_categoria_tabella_id int NULL , " + _
		" 	pub_categoria_field " + SQL_CharField(Conn, 255) + " NULL , " + _
		" 	CONSTRAINT PK_tb_siti_tabelle_pubblicazioni PRIMARY KEY CLUSTERED(pub_id) " + _
		" ) ; " + _
		" ALTER TABLE rel_contents ADD " + _
		"	CONSTRAINT FK_rel_contents__tb_contents__correlato FOREIGN KEY (rcc_correlato_id) REFERENCES tb_contents (co_id), " + _
		"	CONSTRAINT FK_rel_contents__tb_contents__content FOREIGN KEY (rcc_content_id) REFERENCES tb_contents (co_id) " + _
		"		 ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE rel_contents_tags ADD " + _
		"	CONSTRAINT FK_rel_contents_tags__tb_contents FOREIGN KEY (rct_content_id) REFERENCES tb_contents (co_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE , " + _
		"	CONSTRAINT FK_rel_contents_tags__tb_contents_tags FOREIGN KEY (rct_tag_id) REFERENCES tb_contents_tags (tag_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE rel_index_Admin ADD " + _
		"	CONSTRAINT FK_rel_index_admin__tb_admin FOREIGN KEY (ria_admin_id) REFERENCES tb_admin (id_admin) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE, " + _
		"	CONSTRAINT FK_rel_index_admin__tb_contents_index FOREIGN KEY (ria_index_id) REFERENCES tb_contents_index (idx_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_contents ADD " + _
		" 	CONSTRAINT FK_tb_contents__tb_siti_tabelle FOREIGN KEY (co_F_table_id) REFERENCES tb_siti_tabelle (tab_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_contents_index ADD " + _
		"	CONSTRAINT FK_tb_contents_index__tb_contents FOREIGN KEY (idx_content_id) REFERENCES tb_contents (co_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; " + _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	CONSTRAINT FK_tb_siti_tabelle__tb_siti FOREIGN KEY (tab_sito_id) REFERENCES tb_siti (id_sito) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE; " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
		"	CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_contents_index FOREIGN KEY (pub_index_padre_base_id) REFERENCES tb_contents_index (idx_id), " + _
		"	CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_siti_tabelle FOREIGN KEY (pub_tabella_id) REFERENCES tb_siti_tabelle (tab_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE ; "
	 if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__21 =  Aggiornamento__FRAMEWORK_CORE__21 + _
			" ALTER TABLE tb_contents_index ADD " + _
			"	CONSTRAINT FK_tb_contents_index__tb_contents_index__padre FOREIGN KEY (idx_padre_id) REFERENCES tb_contents_index (idx_id), " + _
			"	CONSTRAINT FK_tb_contents_index__tb_contents_index__padre_base FOREIGN KEY (idx_tipologia_padre_base) REFERENCES tb_contents_index (idx_id); " + _
			" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
			"	CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_paginesito FOREIGN KEY (pub_pagina_id) REFERENCES tb_paginesito(id_paginesito), " + _
			" 	CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_siti_tabelle__categoria FOREIGN KEY (pub_categoria_tabella_id) REFERENCES tb_siti_tabelle(tab_id) ; " + _
			" ALTER TABLE tb_contents_index NOCHECK CONSTRAINT FK_tb_contents_index__tb_contents_index__padre; " + _
			" ALTER TABLE tb_contents_index NOCHECK CONSTRAINT FK_tb_contents_index__tb_contents_index__padre_base; " + _
			" ALTER TABLE tb_contents_index NOCHECK CONSTRAINT FK_tb_contents_index__tb_contents_index__padre; " + _
			" ALTER TABLE tb_siti_tabelle_pubblicazioni NOCHECK CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_paginesito; " + _
			" ALTER TABLE tb_siti_tabelle_pubblicazioni NOCHECK CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_siti_tabelle__categoria; "
			
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 22
'...........................................................................................
'NEXT-web 5.0: 	aggiunge campo per la gestione dell'sql di lettura delle tabelle
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__22(conn)
	Select case DB_Type(conn)
		case DB_Access, DB_SQL
			Aggiornamento__FRAMEWORK_CORE__22 = _
				" ALTER TABLE tb_siti_Tabelle ADD " + _
				"	tab_from_sql " + SQL_CharField(Conn, 255) + " NULL; " + _
				" UPDATE tb_siti_tabelle SET tab_from_sql = tab_name; " + _
				" ALTER TABLE tb_siti_Tabelle ALTER COLUMN tab_from_sql " + SQL_CharField(Conn, 255) + " NOT NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 23
'...........................................................................................
'NEXT-web 5.0: 	aggiunge campo e relazione (solo per sql server) 
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__23(conn)
	Aggiornamento__FRAMEWORK_CORE__23 = _
		" ALTER TABLE tb_contents_index ADD " + _
		"	idx_link_pagina_id INT NULL; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__23 =  Aggiornamento__FRAMEWORK_CORE__23 + _
		" ALTER TABLE tb_contents_index ADD " + _
		"	CONSTRAINT FK_tb_contents_index__tb_paginesito FOREIGN KEY (idx_link_pagina_id) REFERENCES tb_paginesito(id_paginesito); " + _
		" ALTER TABLE tb_contents_index NOCHECK CONSTRAINT FK_tb_contents_index__tb_paginesito; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 24
'...........................................................................................
'NEXT-web 5.0: 	rimuove vecchio campo e relativa relazione (Solo per sql) ed aggiunge nuovo 
'				campo e relativa relazione (solo per sql)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__24(conn)
	Aggiornamento__FRAMEWORK_CORE__24 = _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni DROP CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_contents_index ; " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni DROP COLUMN pub_index_padre_base_id ; " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
		"	pub_padre_index_id INT NULL; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__24 =  Aggiornamento__FRAMEWORK_CORE__24 + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
		"	CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_contents_index FOREIGN KEY (pub_padre_index_id) REFERENCES tb_contents_index(idx_id); " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni NOCHECK CONSTRAINT FK_tb_siti_tabelle_pubblicazioni__tb_contents_index; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 25
'...........................................................................................
'NEXT-web 5.0: 	rimuove relazioni tra lingue e pagine mettendola opzionale
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__25(conn)
	Aggiornamento__FRAMEWORK_CORE__25 = _
		" ALTER TABLE tb_pages DROP CONSTRAINT FK_tb_pages_tb_cnt_lingue; " + _
		" ALTER TABLE tb_pages ALTER COLUMN lingua " + replace(SQL_CharField(Conn, 2), "nvarchar", "varchar") + " NULL; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__25 =  Aggiornamento__FRAMEWORK_CORE__25 + _
		" ALTER TABLE tb_pages ADD " + _
		"	CONSTRAINT FK_tb_pages__tb_cnt_lingue FOREIGN KEY (lingua) REFERENCES tb_cnt_lingue(lingua_codice); " + _
		" ALTER TABLE tb_pages NOCHECK CONSTRAINT FK_tb_pages__tb_cnt_lingue; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 26
'...........................................................................................
'NEXT-web 5.0: 	aggiunge campo per gestione colori dei contenuti
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__26(conn)
	Aggiornamento__FRAMEWORK_CORE__26 = _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	tab_colore " + SQL_CharField(Conn, 7) + " NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 27
'...........................................................................................
'NEXT-web 5.0: 	aggiunge campo su descrizione tabelle per gestione url diretto
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__27(conn)
	Aggiornamento__FRAMEWORK_CORE__27 = _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
		"	pub_url_field " + SQL_CharField(Conn, 255) + " NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 28
'...........................................................................................
'NEXT-web 5.0: 	aggiunge tabelle per gestione menu
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__28(conn)
	Aggiornamento__FRAMEWORK_CORE__28 = _
		" CREATE TABLE  " & SQL_Dbo(Conn) & "tb_menu ( " + _
		"	m_id int IDENTITY (1, 1) NOT NULL , " + _
		"	m_id_webs int NOT NULL , " + _
		"	m_nome_it " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	m_nome_en " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	m_nome_fr " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	m_nome_es " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	m_nome_de " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	CONSTRAINT PK_tb_menu PRIMARY KEY CLUSTERED (m_id), " + _
		"	CONSTRAINT FK_tb_menu_tb_webs FOREIGN KEY (m_id_webs) REFERENCES tb_webs (id_webs) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE " + _
		" ) ; " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_menuItem ( " + _
		"	mi_id int IDENTITY (1, 1) NOT NULL , " + _
		"	mi_ordine int NULL , " + _
		"	mi_menu_id int NOT NULL , " + _
		"	mi_index_id int NULL , " + _
		"	mi_attivo bit NOT NULL , " + _
		"	mi_target " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	mi_titolo_it " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_titolo_en " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_titolo_fr " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_titolo_de " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_titolo_es " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_link_it " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_link_en " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_link_fr " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_link_de " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_link_es " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_image_it " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_image_en " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_image_fr " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_image_de " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_image_es " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_tag_title_it " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_tag_title_en " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_tag_title_fr " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_tag_title_de " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	mi_tag_title_es " + SQL_CharField(Conn, 255) + " NULL ," + _
		"	CONSTRAINT PK_tb_menuItem PRIMARY KEY  CLUSTERED (mi_id), " + _
		"	CONSTRAINT FK_tb_menuItem_tb_menu FOREIGN KEY (mi_menu_id) REFERENCES tb_menu (m_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE " + _
		" ) ; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__28 =  Aggiornamento__FRAMEWORK_CORE__28 + _
		" ALTER TABLE tb_menuItem ADD " + _
		"	CONSTRAINT FK_tb_menuItem_tb_contents_index FOREIGN KEY (mi_index_id) REFERENCES tb_contents_index(idx_id); " + _
		" ALTER TABLE tb_menuItem NOCHECK CONSTRAINT FK_tb_menuItem_tb_contents_index; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 29
'...........................................................................................
'NEXT-web 5.0: 	aggiunge relazione tra testata dei menu ed indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__29(conn)
	Aggiornamento__FRAMEWORK_CORE__29 = _
		" ALTER TABLE tb_menu ADD m_index_id INT NULL ; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__29 =  Aggiornamento__FRAMEWORK_CORE__29 + _
		" ALTER TABLE tb_menu ADD " + _
		"	CONSTRAINT FK_tb_menu_tb_contents_index FOREIGN KEY (m_index_id) REFERENCES tb_contents_index(idx_id); " + _
		" ALTER TABLE tb_menu NOCHECK CONSTRAINT FK_tb_menu_tb_contents_index; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 30
'...........................................................................................
'NEXT-web 5.0: 	aggiunge colonna e relazione tra pagine e paginesito
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__30(conn)
	Aggiornamento__FRAMEWORK_CORE__30 = _
		" ALTER TABLE tb_pages ADD id_PaginaSito INT NULL; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__30 =  Aggiornamento__FRAMEWORK_CORE__30 + _
		" UPDATE tb_pages SET tb_pages.id_PaginaSito = tb_pagineSito.id_pagineSito " + _
		"	FROM (tb_pages INNER JOIN tb_pagineSito ON ( tb_pages.id_page = tb_pagineSito.id_pagDyn_IT OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_EN OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_FR OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_DE OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_ES OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_IT OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_EN OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_FR OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_DE OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_ES )) ; " + _
		" ALTER TABLE tb_pages ADD " + _
		"	CONSTRAINT FK_tb_pages_tb_pagineSito FOREIGN KEY (id_PaginaSito) REFERENCES tb_pagineSito(id_pagineSito); " + _
		" ALTER TABLE tb_pages NOCHECK CONSTRAINT FK_tb_pages_tb_pagineSito; "
	else
		Aggiornamento__FRAMEWORK_CORE__30 =  Aggiornamento__FRAMEWORK_CORE__30 + _
		" UPDATE (tb_pages INNER JOIN tb_pagineSito ON ( tb_pages.id_page = tb_pagineSito.id_pagDyn_IT OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_EN OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_FR OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_DE OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagDyn_ES OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_IT OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_EN OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_FR OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_DE OR " + _
													  " tb_pages.id_page = tb_pagineSito.id_pagStage_ES )) " + _
		" SET tb_pages.id_PaginaSito = tb_pagineSito.id_pagineSito ; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 31
'...........................................................................................
'NEXT-web 5.0: 	aggiunge gestione link direttamente associati al contenuto e parametro ed url
'				direttamente associato al contenuto (rimossi da pubblicazioni)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__31(conn)
	Aggiornamento__FRAMEWORK_CORE__31 = _
		" ALTER TABLE tb_contents ADD " + _
		"	co_link_tipo " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_url_IT " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_url_EN " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_url_FR " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_url_DE " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_url_ES " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	co_link_pagina_id INT NULL; " + _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	tab_field_url " + SQL_CharField(Conn, 255) + " NULL , " + _
		"	tab_parametro " + SQL_CharField(Conn, 255) + " NULL; " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni DROP COLUMN pub_url_field; " + _
		" ALTER TABLE tb_siti_tabelle_pubblicazioni DROP COLUMN pub_parametro; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__31 =  Aggiornamento__FRAMEWORK_CORE__31 + _
		" ALTER TABLE tb_contents ADD " + _
		"	CONSTRAINT FK_tb_contents__tb_pagineSito FOREIGN KEY (co_link_pagina_id) REFERENCES tb_pagineSito(id_pagineSito); " + _
		" ALTER TABLE tb_contents NOCHECK CONSTRAINT FK_tb_contents__tb_pagineSito; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 32
'...........................................................................................
'NextCom:	aggiunge campo per archiviazione email su database "ARCHIVIO"
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__32(conn)
	Aggiornamento__FRAMEWORK_CORE__32 = _
		" ALTER TABLE tb_email ADD " + _
		"	email_archiviata BIT NULL, " + _
		" 	email_archiviata_il smalldatetime NULL " + _
		" ; " + _
		" UPDATE tb_email SET email_archiviata = 0 ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 33
'...........................................................................................
'NextWeb 5.0: aggiunge campi immagine ed indice alla pubblicazione.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__33(conn)
	Aggiornamento__FRAMEWORK_CORE__33 = _
		" ALTER TABLE tb_contents_index ADD " + _
		"	idx_ordine INT NULL, " + _
		" 	idx_foto_thumb " + SQL_CharField(Conn, 255) + " NULL, " + _
		" 	idx_foto_zoom " + SQL_CharField(Conn, 255) + " NULL " + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 34
'...........................................................................................
'NextWeb 5.0: aggiunge viste per gestione indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__34(conn)
	Aggiornamento__FRAMEWORK_CORE__34 = _
		" CREATE VIEW " & SQL_Dbo(Conn) & "v_indice AS " + vbCrLf + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"; " + _
		"CREATE VIEW " & SQL_Dbo(Conn) & "v_indice_visibile AS " + vbCrLF + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"    WHERE tb_contents.co_visibile=1 AND " + vbCrLF + _
		"          tb_contents_index.idx_visibile_assoluto=1 AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 35
'...........................................................................................
'NextWeb 5.0: corregge vista "indice visibile"
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__35(conn)
	Aggiornamento__FRAMEWORK_CORE__35 = _
		DropObject(conn, "v_indice_visibile", "VIEW") + _
		"CREATE VIEW " & SQL_Dbo(Conn) & "v_indice_visibile AS " + vbCrLF + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"    WHERE tb_contents.co_visibile=1 AND " + vbCrLF + _
		"          tb_contents_index.idx_visibile_assoluto=1 AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 36
'...........................................................................................
'NextWeb 5.0: aggiunge campo per pubblicazione sito in locale
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__36(conn)
	Aggiornamento__FRAMEWORK_CORE__36 = _
		" ALTER TABLE tb_webs ADD " + _
        "    URL_base " + SQL_CharField(Conn, 255) + " NULL, " + _
        "    URL_secure " + SQL_CharField(Conn, 255) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 37
'...........................................................................................
'NextWeb 5.0: aggiunge campo e relazioni per la gestione dei blocchi 
'             delle pubblicazioni automatiche
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__37(conn)
	Aggiornamento__FRAMEWORK_CORE__37 = _
		" ALTER TABLE tb_contents_index ADD " + _
        "       idx_autopubblicato BIT NULL; " + _
        " UPDATE tb_contents_index SET idx_autopubblicato=0; " + _
        " ALTER TABLE tb_contents_index ALTER COLUMN idx_autopubblicato BIT NOT NULL; " +_
        " CREATE TABLE " & SQL_Dbo(Conn) & "rel_index_pubblicazioni ( " + _
        "       rip_id INT IDENTITY(1,1) NOT NULL, " + _
        "       rip_idx_id INT NOT NULL, " + _
        "       rip_pub_id INT NOT NULL, " + _
		"	CONSTRAINT PK_rel_index_pubblicazioni PRIMARY KEY CLUSTERED (rip_id), " + _
		"	CONSTRAINT FK_rel_index_pubblicazioni__tb_contents_index FOREIGN KEY (rip_idx_id) " + _
        "       REFERENCES tb_contents_index (idx_id) " + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE, " + _
		"	CONSTRAINT FK_rel_index_pubblicazioni__tb_siti_tabelle_pubblicazioni FOREIGN KEY (rip_pub_id) " + _
        "       REFERENCES tb_siti_tabelle_pubblicazioni (pub_id) " + _
		"		ON DELETE NO ACTION  ON UPDATE NO ACTION " + _
		" ) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 38
'...........................................................................................
'ClassCategorie: aggiunge il campo per la gestione della lista degli IDs dei padri
'...........................................................................................
function AggiornamentoSpeciale__FRAMEWORK_CORE__38(DB, rs, version)
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "ptb_categoriegallery", "catC")
    AggiornamentoSpeciale__FRAMEWORK_CORE__38 = "SELECT * FROM AA_versione"
end function


'...........................................................................................
'   funzione utilizzata anche negli altri script di update degli applicativi
'   che hanno strutture di categorizzazione ad albero.
'...........................................................................................
sub AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, tabella, prefisso)
	dim sql
    if ReadCurrentDbVersion(DB.objconn) = ( version - 1 ) then
        'crea colonna su tabella
        sql = "SELECT TOP 1 * FROM "& tabella
        rs.open sql, conn, adOpenDynamic, adLockOptimistic
        if not FieldExists(rs, prefisso &"_tipologie_padre_lista") then
            rs.close
			
			DB.ReSyncTransactionAlways()
			
            sql = "ALTER TABLE "& tabella &" ADD "& prefisso &"_tipologie_padre_lista " + SQL_CharField(DB.objConn, 255) + " NULL"
            CALL DB.objConn.execute(sql, , adExecuteNoRecords)
        else
            rs.close
        end if
        
        'aggiorna dati categorie
		dim categorie
		set categorie = New objCategorie
		with categorie
            set .conn = DB.objConn
			.tabella = tabella
			.prefisso = prefisso
		end with
		
		'per ogni categoria base
		rs.open "SELECT "& prefisso &"_id FROM "& tabella &" WHERE "& prefisso &"_livello = 0", DB.objConn, adOpenForwardOnly, adLockOptimistic
		while not rs.eof
			categorie.operazioni_ricorsive_tipologia(rs(prefisso &"_id"))
			rs.movenext
		wend
		rs.close
		
        categorie.conn = NULL
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 39
'...........................................................................................
'NextWeb 5.0: rinomina colonne per tracciatura modifiche
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__39(conn)
	Aggiornamento__FRAMEWORK_CORE__39 = _
		" ALTER TABLE tb_objects ADD " + _
        "	obj_insData smalldatetime NULL , " + _
        "	obj_insAdmin_id int NULL , " + _
        "	obj_modData smalldatetime NULL , " + _
        "	obj_modAdmin_id int NULL ; " + _
        " ALTER TABLE tb_pages ADD " + _
        "	page_insData smalldatetime NULL , " + _
        "	page_insAdmin_id int NULL , " + _
        "	page_modData smalldatetime NULL , " + _
        "	page_modAdmin_id int NULL ; " + _
        " ALTER TABLE tb_pagineSito ADD " + _
        "	ps_insData smalldatetime NULL , " + _
        "	ps_insAdmin_id int NULL , " + _
        "	ps_modData smalldatetime NULL , " + _
        "	ps_modAdmin_id int NULL ; " + _
        " ALTER TABLE tb_webs ADD " + _
        "	webs_insData smalldatetime NULL , " + _
        "	webs_insAdmin_id int NULL , " + _
        "	webs_modData smalldatetime NULL , " + _
        "	webs_modAdmin_id int NULL ; " + _
        " UPDATE tb_objects SET obj_insData = insData, obj_insAdmin_id=insAdmin_id, obj_modData = modData, obj_modAdmin_id = modAdmin_id ; " + _
        " UPDATE tb_pages SET page_insData = insData, page_insAdmin_id=insAdmin_id, page_modData = modData, page_modAdmin_id = modAdmin_id ; " + _
        " UPDATE tb_pagineSito SET ps_insData = insData, ps_insAdmin_id=insAdmin_id, ps_modData = modData, ps_modAdmin_id = modAdmin_id ; " + _
        " UPDATE tb_webs SET webs_insData = insData, webs_insAdmin_id=insAdmin_id, webs_modData = modData, webs_modAdmin_id = modAdmin_id ; " + _
        " ALTER TABLE tb_objects DROP COLUMN insData, insAdmin_id, modData, modAdmin_id ; " + _
        " ALTER TABLE tb_pages DROP COLUMN insData, insAdmin_id, modData, modAdmin_id ; " + _
        " ALTER TABLE tb_pagineSito DROP COLUMN insData, insAdmin_id, modData, modAdmin_id ; " + _
        " ALTER TABLE tb_webs DROP COLUMN insData, insAdmin_id, modData, modAdmin_id ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 40
'...........................................................................................
'NextWeb 5.0: aggiunge campo per tipizzazione degli oggetti .NET tra i plugin
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__40(conn)
	Aggiornamento__FRAMEWORK_CORE__40 = _
		" ALTER TABLE tb_objects ADD " + _
        "    obj_type " + SQL_CharField(Conn, 255) + " NULL ; " + _
        " UPDATE tb_objects SET obj_type='ascx'; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 41
'...........................................................................................
'NextWeb 5.0: corregge vista "indice visibile" per errore su ACCESS
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__41(conn)
	Aggiornamento__FRAMEWORK_CORE__41 = _
        DropObject(conn, "v_indice_visibile", "VIEW") + _
		"CREATE VIEW " & SQL_Dbo(Conn) & "v_indice_visibile AS " + vbCrLF + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"    WHERE " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 42
'...........................................................................................
'NextWeb 5.0: aggiunge campo per numero di pagina di login dell'area riservata
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__42(conn)
	Aggiornamento__FRAMEWORK_CORE__42 = _
        " ALTER TABLE tb_webs ADD " + _
        "   id_login_page_riservata INT NULL ; "
    if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__42 =  Aggiornamento__FRAMEWORK_CORE__42 + _
		" ALTER TABLE tb_webs ADD " + _
	    "	CONSTRAINT FK_tb_webs_tb_pagineSito_login_riservata FOREIGN KEY (id_login_page_riservata) REFERENCES tb_pagineSito (id_pagineSito) ; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 43
'...........................................................................................
'NextWeb 5.0: aggiunge campo "nome aggiuntivo" per le pagine del nextweb
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__43(conn)
	Aggiornamento__FRAMEWORK_CORE__43 = _
        " ALTER TABLE tb_pagineSito ADD " + _
        "   nome_ps_interno " + SQL_CharField(Conn, 250) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 44
'...........................................................................................
'NextWeb 5.0: aggiunge campo time stamp al sito per la verifica di variazioni ai dati principali
'             del sito o la variazione dell'array delle pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__44(conn)
	Aggiornamento__FRAMEWORK_CORE__44 = _
        " ALTER TABLE tb_webs ADD " + _
        "   webs_modData_pagine SMALLDATETIME NULL; " + _
        " UPDATE tb_webs SET webs_modData_pagine=" + SQL_Now(conn) + ";" + _
        " UPDATE tb_webs SET webs_modData=" + SQL_Now(conn) + " WHERE " + SQL_IsNull(conn, "webs_modData") + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 45
'...........................................................................................
'aggiunge le tabelle del carnet (in teoria utilizzabile anche per NextWeb 4)
'N.B.: in Govenice e AptPortals non inserire perche esiste gia una tabella tb_carnet
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__45(conn)
	Aggiornamento__FRAMEWORK_CORE__45 = _
		" CREATE TABLE " + SQL_Dbo(Conn) + "tb_carnet ( " + vbCrLF + _
		"	car_id " + SQL_PrimaryKey(conn, "tb_carnet") + ", " + vbCrLF + _
		"	car_request " + SQL_CharField(Conn, 0) + " NULL, " + vbCrLF + _
		"	car_dataCreazione SMALLDATETIME NOT NULL, " + vbCrLF + _
		"	cat_ut_id INT NOT NULL" + vbCrLF + _
		" ); " + _
		" CREATE TABLE " + SQL_Dbo(Conn) + "rel_carnet_index (" + vbCrLf + _
		"	rci_id " + SQL_PrimaryKey(conn, "rel_carnet_index") + "," + vbCrLf + _
		"	rci_car_id INT NOT NULL," + vbCrLf + _
		"	rci_idx_id INT NULL," + vbCrLf + _
		" 	rci_url " + SQL_CharField(Conn, 255) + " NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		" ALTER TABLE rel_carnet_index ADD CONSTRAINT FK_rel_carnet_index__tb_carnet " + vbCrLf + _
	    "	FOREIGN KEY (rci_car_id) REFERENCES tb_carnet (car_id) " + vbCrLf + _
        "   ON UPDATE CASCADE ON DELETE CASCADE;"
	'aggiugo le relazioni senza integrita per sql server
	if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__45 = Aggiornamento__FRAMEWORK_CORE__45 + vbCrLF + _
				" ALTER TABLE tb_carnet ADD CONSTRAINT FK_tb_carnet_tb_utenti" + vbCrLf + _
				" 	FOREIGN KEY (cat_ut_id) REFERENCES tb_utenti (ut_id);" + vbCrLf + _
				" ALTER TABLE tb_carnet NOCHECK CONSTRAINT FK_tb_carnet_tb_utenti;" + vbCrLf + _
				" ALTER TABLE rel_carnet_index ADD CONSTRAINT FK_rel_carnet_index_tb_contents_index" + vbCrLf + _
				" 	FOREIGN KEY (rci_idx_id) REFERENCES tb_contents_index (idx_id);" + vbCrLf + _
				" ALTER TABLE rel_carnet_index NOCHECK CONSTRAINT FK_rel_carnet_index_tb_contents_index;"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 46
'...........................................................................................
'aggiorna il carnet con l'ID di sessione ed il path per l'immagine della pagina.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__46(conn)
	Aggiornamento__FRAMEWORK_CORE__46 = _
		" ALTER TABLE tb_carnet " + SQL_AddColumn(conn) + vbCrLf + _
		" 	car_session_id " + SQL_CharField(Conn, 250) + " NULL;" + vbCrLf + _
		" ALTER TABLE rel_carnet_index " + SQL_AddColumn(conn) + vbCrLf + _
		" 	rci_resource_path " + SQL_CharField(Conn, 255) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 47
'...........................................................................................
'aggiorna il carnet aggiungendo il thumbnail
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__47(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__47 = _
		" ALTER TABLE tb_carnet DROP CONSTRAINT FK_tb_carnet_tb_utenti;"
	end if
	
	Aggiornamento__FRAMEWORK_CORE__47 = Aggiornamento__FRAMEWORK_CORE__47 + _
		" ALTER TABLE tb_carnet DROP COLUMN cat_ut_id;" + _
		" ALTER TABLE tb_carnet " + SQL_AddColumn(conn) + _
		" 	car_ut_id INT NULL;" + _
		" ALTER TABLE rel_carnet_index " + SQL_AddColumn(conn) + _
		" 	rci_resourceAlt_path " + SQL_CharField(Conn, 255) + " NULL;"
	
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__47 = Aggiornamento__FRAMEWORK_CORE__47 + _
		" ALTER TABLE tb_carnet ADD CONSTRAINT FK_tb_carnet_tb_utenti" + _
		" 	FOREIGN KEY (car_ut_id) REFERENCES tb_utenti (ut_id);"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 48
'...........................................................................................
'NextWeb 5.0: corregge vista "indice visibile" per controllo date comprese nell'intervallo
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__48(conn)
	Aggiornamento__FRAMEWORK_CORE__48 = _
        DropObject(conn, "v_indice_visibile", "VIEW") + _
		"CREATE VIEW " & SQL_Dbo(Conn) & "v_indice_visibile AS " + vbCrLF + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"    WHERE " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 49
'...........................................................................................
'aggiunge settaggi interni all'editor per il sito
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__49(conn)
	Aggiornamento__FRAMEWORK_CORE__49 = _
        " ALTER TABLE tb_webs " + SQL_AddColumn(conn) + _
        "   sito_accessibile BIT NULL, " + _
        "   editor_guide_visibili BIT NULL, " + _
        "   editor_guide_colore " + SQL_CharField(Conn, 7) + " NULL, " + _
        "   editor_guide_posizioni_visibili BIT NULL, " + _
        "   editor_help_attivo BIT NULL " + _
        " ; " + _
        " UPDATE tb_webs SET sito_accessibile = 1, " + _
                           " editor_guide_visibili = 1, " + _
                           " editor_guide_colore='#000000', " + _
                           " editor_guide_posizioni_visibili = 1, " + _
                           " editor_help_attivo = 0 " + _
        " ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 50
'...........................................................................................
'NEXT-web 5.0: 	corregge problema di creazione del campo lingua delle pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__50(conn)
    if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__50 = " ALTER TABLE tb_pages DROP CONSTRAINT FK_tb_pages__tb_cnt_lingue; "
    end if
    Aggiornamento__FRAMEWORK_CORE__50 = Aggiornamento__FRAMEWORK_CORE__50 + _
        " ALTER TABLE tb_pages ADD lingua_tmp " + replace(SQL_CharField(Conn, 2), "nvarchar", "varchar") + " NULL; " + _
        " UPDATE tb_pages SET lingua_tmp = lingua; " + _
        " ALTER TABLE tb_pages DROP COLUMN lingua; " + _
        " ALTER TABLE tb_pages ADD lingua " + replace(SQL_CharField(Conn, 2), "nvarchar", "varchar") + " NULL; " + _
        " UPDATE tb_pages SET lingua = lingua_tmp; "
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__50 =  Aggiornamento__FRAMEWORK_CORE__50 + _
		" ALTER TABLE tb_pages ADD " + _
		"	CONSTRAINT FK_tb_pages__tb_cnt_lingue FOREIGN KEY (lingua) REFERENCES tb_cnt_lingue(lingua_codice); " + _
		" ALTER TABLE tb_pages NOCHECK CONSTRAINT FK_tb_pages__tb_cnt_lingue; "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 51
'...........................................................................................
'aggiunge campo foto alle categorie di gallery
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__51(conn)
    Aggiornamento__FRAMEWORK_CORE__51 = _
        " ALTER TABLE ptb_categoriegallery ADD catC_foto " + SQL_CharField(Conn, 250) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 52
'...........................................................................................
'aggiunge campo valutazione su tb_comments
'annullato per problemi di aggiornamenti del next-comments
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__52(conn)
    Aggiornamento__FRAMEWORK_CORE__52 = _
        " SELECT * FROM AA_versione;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 53
'...........................................................................................
'corregge dimensione campi nome delle categorie di news e di link
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__53(conn)
    Aggiornamento__FRAMEWORK_CORE__53 = _
        " ALTER TABLE tb_news_categorie ALTER COLUMN " + "cat_nome_it " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_news_categorie ALTER COLUMN " + "cat_nome_en " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_news_categorie ALTER COLUMN " + "cat_nome_fr " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_news_categorie ALTER COLUMN " + "cat_nome_de " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_news_categorie ALTER COLUMN " + "cat_nome_es " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_links_categorie ALTER COLUMN " + "cat_nome_it " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_links_categorie ALTER COLUMN " + "cat_nome_en " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_links_categorie ALTER COLUMN " + "cat_nome_fr " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_links_categorie ALTER COLUMN " + "cat_nome_de " + SQL_CharField(Conn, 255) + " NULL; " + _
        " ALTER TABLE tb_links_categorie ALTER COLUMN " + "cat_nome_es " + SQL_CharField(Conn, 255) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 54
'...........................................................................................
'elimina i commenti che vengono da ora gestiti tramite un file di aggiornamento separato
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__54(conn)
	Aggiornamento__FRAMEWORK_CORE__54 = _
        DropObject(conn, "tb_comments", "TABLE")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 55
'...........................................................................................
'aggiunge un campo tipo alle voci del carnet
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__55(conn)
	Aggiornamento__FRAMEWORK_CORE__55 = _
		" ALTER TABLE rel_carnet_index ADD rci_tipo INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 56
'...........................................................................................
'aggiunge i formati delle immagini al NextWeb5
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__56(conn)
	Aggiornamento__FRAMEWORK_CORE__56 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_immaginiFormati (" + vbCrLf + _
		"	imf_id "& SQL_PrimaryKey(conn, "tb_immaginiFormati") + ", " + vbCrLf + _
		"	imf_webId INT NULL, " + vbCrLf + _
		"	imf_nome " + SQL_CharField(Conn, 255) + ", " + vbCrLf + _
		"	imf_thumb_width INT NULL, " + vbCrLf + _
		"	imf_thumb_height INT NULL, " + vbCrLf + _
		"	imf_thumb_dir " + SQL_CharField(Conn, 100) + " NULL, " + vbCrLf + _
		"	imf_zoom_width INT NULL, " + vbCrLf + _
		"	imf_zoom_height INT NULL, " + vbCrLf + _
		"	imf_zoom_dir " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLf + _
		" 	imf_suffisso BIT NULL, " + vbCrLf + _
		AddInsModFields("imf") + _
		" );" + vbCrLf + _
		AddInsModRelations(conn, "tb_immaginiFormati", "imf") + _
		SQL_AddForeignKey(conn, "tb_immaginiFormati", "imf_webId", "tb_webs", "id_webs", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 57
'...........................................................................................
'modifica i formati delle immagini del NextWeb5
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__57(conn)
	Aggiornamento__FRAMEWORK_CORE__57 = _
		" ALTER TABLE tb_immaginiFormati DROP COLUMN" + vbCrLf + _
		" 	imf_thumb_width, imf_thumb_height, imf_thumb_dir, imf_zoom_width, imf_zoom_height, imf_zoom_dir, imf_suffisso;" + vbCrLf + _
		" ALTER TABLE tb_immaginiFormati ADD" + vbCrLf + _
		"	imf_width INT NULL, " + vbCrLf + _
		"	imf_height INT NULL, " + vbCrLf + _
		"	imf_suffisso "+ SQL_CharField(Conn, 50) +" NULL, " + vbCrLf + _
		"	imf_suffissoFormato BIT NULL, " + vbCrLf + _
		"	imf_dir "+ SQL_CharField(Conn, 255) +" NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 58
'...........................................................................................
'aggiunge flag per disabilitare la visualizzazione dei figli del menu
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__58(conn)
	Aggiornamento__FRAMEWORK_CORE__58 = _
		" ALTER TABLE tb_menuItem ADD" + vbCrLf + _
		"	mi_figli BIT NULL;" + vbCrLf + _
		" UPDATE tb_menuItem SET mi_figli = 1"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 59
'...........................................................................................
'toglie integrità al carnet che e stata rimessa dopo un aggiornamento errato
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__59(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__59 = "ALTER TABLE tb_carnet NOCHECK CONSTRAINT FK_tb_carnet_tb_utenti"
	else
		Aggiornamento__FRAMEWORK_CORE__59 = "SELECT * FROM aa_versione"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 60
'...........................................................................................
'flag su template per visualizzazione semplificata (per e-mail)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__60(conn)
	Aggiornamento__FRAMEWORK_CORE__60 = _
		"ALTER TABLE tb_pages ADD" + vbCrLf + _
		"	semplificata BIT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 61
'...........................................................................................
'aggiunge contatori accessi per indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__61(conn)
	Aggiornamento__FRAMEWORK_CORE__61 = _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ADD" + vbCrLf + _
		"	idx_contatore INT NULL," + vbCrLf + _
		"	idx_contUtenti INT NULL," + vbCrLf + _
		"	idx_contCrawler INT NULL," + vbCrLf + _
		"	idx_contAltro INT NULL," + vbCrLf + _
		"	idx_contRes SMALLDATETIME NULL;" + vbCrLf + _
		"UPDATE tb_contents_index SET idx_contatore = 0, idx_contUtenti = 0, idx_contCrawler = 0, idx_contAltro = 0"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 62
'...........................................................................................
'aggiunge tabella storico per indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__62(conn)
	Aggiornamento__FRAMEWORK_CORE__62 = _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_storico_index (" + vbCrLf + _
		"	si_id " & SQL_PrimaryKey(conn, "tb_storico_index") + ", " + vbCrLf + _
		"	si_sw_id INT NULL," + vbCrLf + _
		"	si_idx_id INT NULL," + vbCrLf + _
		"	si_idx_padre_id INT NULL," + vbCrLf + _
		"	si_idx_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	si_idx_foglia BIT NULL," + vbCrLf + _
		"	si_idx_livello INT NULL," + vbCrLf + _
		"	si_tab_id INT NULL," + vbCrLf + _
		"	si_tab_name " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
			SQL_MultiLanguageField("si_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	si_link_pagina_id INT NULL," + vbCrLf + _
			SQL_MultiLanguageField("si_link_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	si_contatore INT NULL," + vbCrLf + _
		"	si_contUtenti INT NULL," + vbCrLf + _
		"	si_contCrawler INT NULL," + vbCrLf + _
		"	si_contAltro INT NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		" ALTER TABLE " & SQL_Dbo(Conn) & "tb_storico_index ADD CONSTRAINT FK_tb_storico_index_tb_storico_webs " + vbCrLf + _
		"	FOREIGN KEY (si_sw_id) REFERENCES tb_storico_webs (sw_ID) " + vbCrLf + _
		"	ON DELETE CASCADE  ON UPDATE CASCADE ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 63
'...........................................................................................
'imposta contatori degli accessi per indice come non null
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__63(conn)
	Aggiornamento__FRAMEWORK_CORE__63 = _
		"UPDATE tb_contents_index SET idx_contatore = 0, idx_contUtenti = 0, idx_contCrawler = 0, idx_contAltro = 0, idx_contRes = " & SQL_NOW(conn) & + "; " + vbCrLF + _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ALTER COLUMN idx_contatore INT NOT NULL; " + vbCrLf + _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ALTER COLUMN idx_contUtenti INT NOT NULL; " + vbCrLf + _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ALTER COLUMN idx_contCrawler INT NOT NULL; " + vbCrLf + _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ALTER COLUMN idx_contAltro INT NOT NULL; " + vbCrLf + _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_contents_index ALTER COLUMN idx_contRes SMALLDATETIME NOT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 64
'...........................................................................................
'aggiunge la chiave esterna allo storico dell'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__64(conn)
	Aggiornamento__FRAMEWORK_CORE__64 = _
		"ALTER TABLE " & SQL_Dbo(Conn) & "tb_storico_index ADD si_co_F_key_id INT NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 65
'...........................................................................................
'aggiunge chiavi di aggancio a Google Analytics e Google Maps
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__65(conn)
	Aggiornamento__FRAMEWORK_CORE__65 = _
		"ALTER TABLE tb_webs ADD " + _
		"	google_analytics_code " + SQL_CharField(Conn, 255) + " NULL" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 66
'...........................................................................................
'aggiunge latitudine e longitudine derivati da google maps.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__66(conn)
	Aggiornamento__FRAMEWORK_CORE__66 = _
		"ALTER TABLE tb_Indirizzario ADD " + _
		"	google_maps_latitudine REAL NULL, " + vbCrLf + _
		"	google_maps_longitudine REAL NULL " + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 67
'...........................................................................................
'aggiunge tabella di configurazione per google maps
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__67(conn)
	Aggiornamento__FRAMEWORK_CORE__67 = _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_webs_directories (" + vbCrLf + _
		"	dir_id " & SQL_PrimaryKey(conn, "tb_webs_directories") + ", " + vbCrLf + _ 
		"	dir_web_id INT NOT NULL, " + vbCrLf + _
		"	dir_url " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLf + _
		"	dir_google_maps_key " + SQL_CharField(Conn, 255) + " NULL " + vbCrLf + _
		" ); " + vbCrLF + _
		SQL_AddForeignKey(conn, "tb_webs_directories", "dir_web_id", "tb_webs", "id_webs", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 68
'...........................................................................................
'corregge tipo di dato per latitudine e longitudine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__68(conn)
	Aggiornamento__FRAMEWORK_CORE__68 = _
		"ALTER TABLE tb_Indirizzario ALTER COLUMN google_maps_latitudine FLOAT NULL ; " + vbCrLf + _
		"ALTER TABLE tb_Indirizzario ALTER COLUMN google_maps_longitudine FLOAT NULL ; " + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 69
'...........................................................................................
'crea vista globale per i contatti del next-com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__69(conn)
	Aggiornamento__FRAMEWORK_CORE__69 = _
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indirizzario AS " + _
		"	SELECT *," + vbCrLF + _
		"		(SELECT TOP 1 valoreNumero FROM tb_valoriNumeri WHERE id_indirizzario = IDElencoIndirizzi AND id_tipoNumero = " & VAL_EMAIL & " ORDER BY email_default DESC ) AS email," + vbCrLF + _
		"		(SELECT TOP 1 valoreNumero FROM tb_valoriNumeri WHERE id_indirizzario = IDElencoIndirizzi AND id_tipoNumero = " & VAL_TELEFONO & ") AS telefono," + vbCrLf + _
		"		(SELECT TOP 1 valoreNumero FROM tb_valoriNumeri WHERE id_indirizzario = IDElencoIndirizzi AND id_tipoNumero = " & VAL_FAX & ") AS fax, " + vbCrLf + _
		"		(SELECT TOP 1 valoreNumero FROM tb_valoriNumeri WHERE id_indirizzario = IDElencoIndirizzi AND id_tipoNumero = " & VAL_CELLULARE & ") AS cellulare," + vbCrLf + _
		"		(SELECT TOP 1 valoreNumero FROM tb_valoriNumeri WHERE id_indirizzario = IDElencoIndirizzi AND id_tipoNumero = " & VAL_URL & ") AS sitoWeb " + vbCrLF + _
		"	FROM tb_indirizzario " + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 70
'...........................................................................................
'aggiunge chiavi di verifica a google webmaster tools
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__70(conn)
	Aggiornamento__FRAMEWORK_CORE__70 = _
		"ALTER TABLE tb_webs ADD " + _
		"	google_webmaster_tools_verify_code " + SQL_CharField(Conn, 255) + " NULL" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 71
'...........................................................................................
'aggiunge vista per generazione sitemap
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__71(conn)
	Aggiornamento__FRAMEWORK_CORE__71 = _
		" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
		"    SELECT DISTINCT URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_IT AS URI, idx_modData, id_webs " + vbCrLF + _
		"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
		"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
		"    WHERE (idx_link_url_IT <> '') "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 72
'...........................................................................................
'aggiunge cartella xml a nextweb5 dentro upload e ripulisce directory temp
'...........................................................................................
function AggiornamentoSpeciale__FRAMEWORK_CORE__72(DB, version)
	dim sql, fso, FolderUpload, FolderSite
	'esegue un aggiornamento fasullo per aumentare la versione
	sql = "SELECT * FROM AA_Versione"
	CALL DB.Execute(sql, version)
	if DB.last_update_executed then
		'esegue aggiornamento solo se la versione e' corretta e la query fasulla e' stata eseguita
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set FolderUpload = fso.GetFolder(Application("IMAGE_PATH"))
		
		'ripulisce directory temporanee
		CALL ClearTempDir(fso)
		
		'rimuove file inutili dalle directory (qualsiasi directory)
		CALL FileRemove(fso, Application("IMAGE_PATH"), "thumbs.db", true)
		CALL FileRemove(fso, Application("IMAGE_PATH"), "pspbrwse.jbf", true)
		
		if GetNextWebCurrentVersion(conn, rs) > 4 then
			'scorre tutte le directory dei siti (solo con nome numerico)
			for each FolderSite in FolderUpload.SubFolders
				if isNumeric(FolderSite.name) then
					if not fso.FolderExists(FolderSite.path + "\xml") then
						CALL fso.CreateFolder(FolderSite.path + "\xml")
					end if
				end if
			next
		end if
		
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 73
'...........................................................................................
'corregge errore di dichiarazione raggruppamenti nelle tabelle dell'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__73(conn)
	Aggiornamento__FRAMEWORK_CORE__73 = _
		" UPDATE tb_siti_tabelle SET tab_field_chiave = 'co_id' WHERE tab_name LIKE 'tb_contents_index' "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 74
'...........................................................................................
'aggiunge meta tag all'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__74(conn)
	Aggiornamento__FRAMEWORK_CORE__74 = _
		" ALTER TABLE tb_contents ADD" + vbCrLf + _
		"	co_meta_keywords_it " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_keywords_en " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_keywords_fr " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_keywords_es " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_keywords_de " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_description_it " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_description_en " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_description_fr " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_description_es " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_meta_description_de " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	co_alt_it " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	co_alt_en " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	co_alt_fr " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	co_alt_es " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	co_alt_de " + SQL_CharField(Conn, 255) + ";" + vbCrLf + _
		" ALTER TABLE tb_contents_index ADD" + vbCrLf + _
		"	idx_meta_keywords_it " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_keywords_en " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_keywords_fr " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_keywords_es " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_keywords_de " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_description_it " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_description_en " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_description_fr " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_description_es " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_meta_description_de " + SQL_CharField(Conn, 0) + "," + vbCrLf + _
		"	idx_alt_it " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	idx_alt_en " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	idx_alt_fr " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	idx_alt_es " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	idx_alt_de " + SQL_CharField(Conn, 255) + ";" + vbCrLf + _
		" ALTER TABLE tb_siti_tabelle ADD" + vbCrLf + _
		"	tab_field_meta_keywords" + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	tab_field_meta_description" + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	tab_field_alt" + SQL_CharField(Conn, 255) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 75
'...........................................................................................
'modifica dichiarazione campi meta tag su tb_siti_Tabelle
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__75(conn)
	Aggiornamento__FRAMEWORK_CORE__75 = _
		" ALTER TABLE tb_siti_tabelle ALTER COLUMN tab_field_meta_keywords" + SQL_CharField(Conn, 0) + "; " + vbCrLf + _
		" ALTER TABLE tb_siti_tabelle ALTER COLUMN tab_field_meta_description" + SQL_CharField(Conn, 0) + "; " + vbCrLf + _
		" ALTER TABLE tb_siti_tabelle ALTER COLUMN tab_field_alt" + SQL_CharField(Conn, 0) + "; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 76
'...........................................................................................
'corregge descrizione tabelle per indice.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__76(conn)
	Aggiornamento__FRAMEWORK_CORE__76 = _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_descrizione = 'META_description_', " + _
		"	tab_field_meta_keywords='META_keywords_', " + _
		"	tab_field_meta_description='META_description_', " + _
		"	tab_field_alt='titolo_' " + _
		" WHERE tab_name LIKE 'tb_webs' ; " + _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_meta_keywords='PAGE_keywords_', " + _
		"	tab_field_meta_description='PAGE_description_' " + _
		" WHERE tab_name LIKE 'tb_pagineSito' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 76
'...........................................................................................
'corregge descrizione pagine per sincronizzazione con l'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__77(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__77 = _
			" UPDATE tb_pagineSito " + _
			"	SET PAGE_description_IT = COALESCE(PAGE_description_IT, co_descrizione_IT), " + _
			"		PAGE_description_EN = COALESCE(PAGE_description_EN, co_descrizione_EN), " + _
			"		PAGE_description_FR = COALESCE(PAGE_description_FR, co_descrizione_FR), " + _
			"		PAGE_description_DE = COALESCE(PAGE_description_DE, co_descrizione_DE), " + _
			"		PAGE_description_ES = COALESCE(PAGE_description_ES, co_descrizione_ES) " + _
    		"	FROM (tb_contents INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id AND tb_siti_tabelle.tab_name LIKE 'tb_paginesito') " + _
			"		 INNER JOIN tb_pagineSito ON tb_contents.co_F_key_id = tb_pagineSito.id_pagineSito ; "
	else
		Aggiornamento__FRAMEWORK_CORE__77 = _
			" UPDATE (( tb_pagineSito INNER JOIN tb_contents ON tb_contents.co_F_key_id = tb_pagineSito.id_pagineSito ) " + _
			" 	     INNER JOIN tb_siti_tabelle ON (tb_contents.co_F_table_id = tb_siti_tabelle.tab_id AND tb_siti_tabelle.tab_name LIKE 'tb_paginesito')) " + _
			"	SET PAGE_description_IT = IIF(PAGE_description_IT <> '' AND NOT IsNull(PAGE_description_IT), PAGE_description_IT, co_descrizione_IT), " + _
			"		PAGE_description_EN = IIF(PAGE_description_EN <> '' AND NOT IsNull(PAGE_description_EN), PAGE_description_EN, co_descrizione_EN), " + _
			"		PAGE_description_FR = IIF(PAGE_description_FR <> '' AND NOT IsNull(PAGE_description_FR), PAGE_description_FR, co_descrizione_FR), " + _
			"		PAGE_description_DE = IIF(PAGE_description_DE <> '' AND NOT IsNull(PAGE_description_DE), PAGE_description_DE, co_descrizione_DE), " + _
			"		PAGE_description_ES = IIF(PAGE_description_ES <> '' AND NOT IsNull(PAGE_description_ES), PAGE_description_ES, co_descrizione_ES) ; "
	end if
	Aggiornamento__FRAMEWORK_CORE__77 = Aggiornamento__FRAMEWORK_CORE__77 + _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_alt = 'nome_ps_', " + _
		"	tab_field_descrizione = 'PAGE_description_', " + _
		"	tab_field_meta_keywords='PAGE_keywords_', " + _
		"	tab_field_meta_description='PAGE_description_' " + _
		" WHERE tab_name LIKE 'tb_pagineSito' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 78
'...........................................................................................
'corregge vista per generazione sitemap
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__78(conn)
	Aggiornamento__FRAMEWORK_CORE__78 = _
		DropObject(conn, "v_indice_sitemap", "VIEW") + _
		" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
		"    SELECT DISTINCT URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_IT AS URI, " + vbCrLF + _
		"    idx_modData, id_webs, idx_id, idx_livello " + vbCrLF + _
		"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
		"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
		"    WHERE (idx_link_url_IT <> '') "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 79
'...........................................................................................
'cancella doppioni generati da errore di salvataggio dei contatti
'per access viene eseguito con l'aggiornamento 331 di Update_A_dbContent.asp
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__79(conn)
	if DB_Type(conn) = DB_ACCESS then
		'per access viene eseguito con l'aggiornamento 331 di Update_A_dbContent.asp
		Aggiornamento__FRAMEWORK_CORE__79 = "SELECT * FROM aa_versione"
	else
		Aggiornamento__FRAMEWORK_CORE__79 = _
			" DELETE FROM rel_rub_ind" + vbCrLf + _
			" WHERE id_rub_ind <>" + vbCrLf + _
		    "	(SELECT MAX(id_rub_ind) FROM rel_rub_ind r" + vbCrLf + _
		    "	 WHERE id_indirizzo = rel_rub_ind.id_indirizzo" + vbCrLf + _
			"	 AND id_rubrica = rel_rub_ind.id_rubrica)"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 80
'...........................................................................................
'genera aggiornamenti per la revisione delle email del next-com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__80(conn)
	Aggiornamento__FRAMEWORK_CORE__80 = _
		" ALTER TABLE log_cnt_email " + SQL_AddColumn(conn) + _
		" 	log_cnt_nominativo " + SQL_CharField(Conn, 255) + " NULL, " + _
		" 	log_inviato_ok BIT NULL " + _
		" ; " + _
		" ALTER TABLE log_cnt_email DROP CONSTRAINT " & IIF(DB_Type(conn) = DB_SQL, "FK_log_cnt_email_tb_Indirizzario", "FK_log_cnt_email__tb_indirizzario") + _
		" ; " + _
		SQL_AddForeignKey(conn, "log_cnt_email", "log_cnt_id", "tb_Indirizzario", "IdElencoIndirizzi", false, "") + _
		" ALTER TABLE tb_email " + SQL_AddColumn(conn) + _
		"	email_isBozza BIT NULL " + _
		" ; " + _
		" CREATE TABLE " & SQL_dbo(conn) & "log_rubriche_email ( " + _
		"	log_id " + SQL_PrimaryKey(conn, "log_rubriche_email") + ", " + _
		"	log_rubrica_id INT NOT NULL, " + _
		"	log_email_id INT NOT NULL " + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "log_rubriche_email", "log_email_id", "tb_email", "email_id", true, "") + _
		SQL_AddForeignKey(conn, "log_rubriche_email", "log_rubrica_id", "tb_rubriche", "id_Rubrica", false, "")
		
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__80 = Aggiornamento__FRAMEWORK_CORE__80 + _
			" UPDATE log_cnt_email " + _
			"	SET log_cnt_nominativo = LEFT(CASE WHEN IsNull(isSocieta, 0) = 1 THEN " + _
			"									CASE WHEN IsNull(CognomeElencoIndirizzi, '')<>'' THEN " + _
			"										NomeOrganizzazioneElencoIndirizzi + ' - ' + CognomeElencoIndirizzi + ' '  + NomeElencoIndirizzi " + _
			"									ELSE " + _
			"										NomeOrganizzazioneElencoIndirizzi " + _
			"									END " + _
			"							 	  ELSE " + _
			"									CASE WHEN IsNull(NomeOrganizzazioneElencoIndirizzi, '')<>'' THEN " + _
			"										CognomeElencoIndirizzi + ' '  + NomeElencoIndirizzi + ' - ' + NomeOrganizzazioneElencoIndirizzi " + _
			"									ELSE " + _
			"										CognomeElencoIndirizzi + ' '  + NomeElencoIndirizzi " + _
			"									END " + _
			"								  END, 255 ) , " + _
			"	log_inviato_ok = 1 " + _
    		"	FROM log_cnt_email INNER JOIN tb_indirizzario ON log_cnt_email.log_cnt_id = tb_indirizzario.idelencoindirizzi ; "
	else
		Aggiornamento__FRAMEWORK_CORE__80 = Aggiornamento__FRAMEWORK_CORE__80 +_
			" UPDATE log_cnt_email INNER JOIN tb_indirizzario ON log_cnt_email.log_cnt_id = tb_indirizzario.idelencoindirizzi " + _
			"	SET log_cnt_nominativo = IIF(isSocieta, " + _
			"									IIF(CognomeElencoIndirizzi<>'', NomeOrganizzazioneElencoIndirizzi & ' - ' & CognomeElencoIndirizzi & ' '  & NomeElencoIndirizzi, NomeOrganizzazioneElencoIndirizzi), " + _
			"									IIF(NomeOrganizzazioneElencoIndirizzi<>'', CognomeElencoIndirizzi & ' '  & NomeElencoIndirizzi & ' - ' & NomeOrganizzazioneElencoIndirizzi, CognomeElencoIndirizzi & ' '  & NomeElencoIndirizzi)), " + _
			"	log_inviato_ok = 1 "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 81
'...........................................................................................
'genera aggiornamenti per la revisione delle email del next-com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__81(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__81 = _
			" UPDATE log_cnt_email " + _
			"	SET log_email = tb_ValoriNumeri.ValoreNumero " + _
    		"	FROM (log_cnt_email INNER JOIN tb_Indirizzario ON log_cnt_email.log_cnt_id = tb_Indirizzario.IDElencoIndirizzi) " + _
			"		 INNER JOIN tb_ValoriNumeri ON (tb_Indirizzario.IDElencoIndirizzi = tb_ValoriNumeri.id_Indirizzario AND tb_ValoriNumeri.id_TipoNumero=" & VAL_EMAIL & " AND  tb_ValoriNumeri.email_default=1 ) " + _
			"   WHERE NOT (log_email LIKE '%@%') "
	else
		Aggiornamento__FRAMEWORK_CORE__81 = _
			" UPDATE (log_cnt_email INNER JOIN tb_Indirizzario ON log_cnt_email.log_cnt_id = tb_Indirizzario.IDElencoIndirizzi) " + _
			"		 INNER JOIN tb_ValoriNumeri ON (tb_Indirizzario.IDElencoIndirizzi = tb_ValoriNumeri.id_Indirizzario AND tb_ValoriNumeri.id_TipoNumero=" & VAL_EMAIL & " AND  tb_ValoriNumeri.email_default ) " + _
			"	SET log_email = tb_ValoriNumeri.ValoreNumero " + _
			"	WHERE NOT (log_email LIKE '%@%') "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 82 (09/05/2008)
'...........................................................................................
'rimuove campi aggiuntivi sulle email non piu' utilizzate e gestione posta in arrivo
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__82(conn)
	Aggiornamento__FRAMEWORK_CORE__82 = _
		" ALTER TABLE tb_email DROP COLUMN email_page_owned; " + _
		" ALTER TABLE tb_email DROP COLUMN email_in; " + _
		" ALTER TABLE tb_email DROP COLUMN email_account; " + _
		" ALTER TABLE tb_email DROP COLUMN email_UIDL; " + _
		" ALTER TABLE tb_email DROP COLUMN email_MessageID; " + _
		" ALTER TABLE tb_email DROP COLUMN email_to; " + _
		" ALTER TABLE tb_email DROP COLUMN email_cc; " + _
		" ALTER TABLE tb_email DROP COLUMN email_from; " & _
		" ALTER TABLE rel_dip_email DROP CONSTRAINT " & IIF(DB_Type(conn) = DB_SQL, "FK_rel_dip_email_tb_admin", "FK_rel_dip_email__tb_admin") + " ; "
	if lcase(GetDatabaseName(conn))<> "aptbibione" AND _
	   lcase(GetDatabaseName(conn))<> "aptchioggia" AND _
	   lcase(GetDatabaseName(conn))<> "aptjesolo" AND _
	   lcase(GetDatabaseName(conn))<> "govenice" then
	   	Aggiornamento__FRAMEWORK_CORE__82 = Aggiornamento__FRAMEWORK_CORE__82 + _
			IIF(DB_Type(conn) = DB_SQL, " ALTER TABLE rel_dip_email DROP CONSTRAINT FK_rel_dip_email_tb_admin1; ", "")
	end if
   	Aggiornamento__FRAMEWORK_CORE__82 = Aggiornamento__FRAMEWORK_CORE__82 + _
		" ALTER TABLE rel_dip_email DROP CONSTRAINT " & IIF(DB_Type(conn) = DB_SQL, "FK_rel_dip_email_tb_email", "FK_rel_dip_email__tb_email") + "  ; " + _
		DropObject(conn, "rel_dip_email", "TABLE") + _
		" ALTER TABLE tb_emailConfig DROP CONSTRAINT " & IIF(DB_Type(conn) = DB_SQL, "FK_tb_emailConfig_tb_admin", "FK_tb_emailConfig__tb_admin") + "; " + _
		DropObject(conn, "tb_emailConfig", "TABLE")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 83
'...........................................................................................
'aggiunge nome rubrica su log di spedizione email
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__83(conn)
	Aggiornamento__FRAMEWORK_CORE__83 = _
		" ALTER TABLE log_rubriche_email " + SQL_AddColumn(conn) + _
		" 	log_rubrica_nome " + SQL_CharField(Conn, 255) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 84
'...........................................................................................
'imposta stato delle email
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__84(conn)
	Aggiornamento__FRAMEWORK_CORE__84 = _
		" UPDATE tb_email SET email_isBozza=0"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 85
'...........................................................................................
'crea tabella per gestione dei filtri di esclusione dei conteggi
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__85(conn)
	Aggiornamento__FRAMEWORK_CORE__85 = _
            " CREATE TABLE " + SQL_Dbo(Conn) + "tb_contents_log_filtri ( " + _
            "   fil_id " + SQL_PrimaryKey(conn, "tb_contents_log_filtri") + ", " + _
            "   fil_parametro " + SQL_CharField(Conn, 50) + " NOT NULL, " + _
            "   fil_valore " + SQL_CharField(Conn, 255) + " NULL," + _
			"	fil_tipo INT NULL" + _
			" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 86
'...........................................................................................
'modifiche all'indice per url rewriting
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__86(conn)
	Aggiornamento__FRAMEWORK_CORE__86 = _
            " ALTER TABLE tb_contents_index ADD " + _
            "   idx_principale BIT NULL," + _
			SQL_MultiLanguageField("idx_link_url_rw_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + _
			" ; " + _
			" ALTER TABLE tb_contents ADD " + _
            SQL_MultiLanguageField("co_link_url_rw_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + _
			" ; " + _
			" ALTER TABLE tb_siti_tabelle_pubblicazioni ADD " + _
			"	pub_field_principale " + SQL_CharField(Conn, 255) + " NULL " + _
			" ; " + _
			" ALTER TABLE tb_siti_tabelle ADD " + _
            SQL_MultiLanguageField("tab_field_titolo_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("tab_field_titolo_alt_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
            SQL_MultiLanguageField("tab_field_descrizione_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
            SQL_MultiLanguageField("tab_field_codice_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
            SQL_MultiLanguageField("tab_field_url_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("tab_field_meta_keywords_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			SQL_MultiLanguageField("tab_field_meta_description_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
			" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 87
'...........................................................................................
'corregge dati su descrizione tabelle
'...........................................................................................
function AggiornamentoSpeciale__FRAMEWORK_CORE__87(DB, version)
	dim sql, rs, field, dFieldName, lingua
	'esegue un aggiornamento fasullo per aumentare la versione
	sql = "SELECT * FROM AA_Versione"
	CALL DB.Execute(sql, version)
	if DB.last_update_executed then
		
		set rs = server.CreateObject("ADODB.recordset")
		sql = "SELECT * FROM tb_siti_tabelle "
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		
		while not rs.eof
			for each field in rs.fields
				if cString(field.value)<>"" then
					Select case lcase(field.name)
						case "tab_field_titolo", _
							 "tab_field_descrizione", _
							 "tab_field_url", _
							 "tab_field_meta_keywords", _
							 "tab_field_meta_description", _
							 "tab_field_alt"
							if lcase(field.name) = "tab_field_alt" then
								dFieldName = "tab_field_titolo_alt"
							else
								dFieldName = field.name
							end if
							
							if instr(1, cString(field.value & " "), "_ ", vbTextCompare) then
								'campo multilingua
								for each lingua in Application("LINGUE")
									rs(dFieldName + "_" + lingua) = left(Trim(Replace(CString(field.value) + " ", "_ ", "_"& lingua &" ")), rs(dFieldName + "_" + lingua).DefinedSize)
								next
							else
								'campo semplice: copia su versione italiana
								rs(dFieldName + "_" + LINGUA_ITALIANO) = field.value
							end if
					end select
				end if
			next
			rs.update
			rs.movenext
		wend
		
		rs.close
		set rs = nothing
    end if
	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 88
'...........................................................................................
'rimuove campi dell'indice non piu' utilizzati
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__88(conn)
	Aggiornamento__FRAMEWORK_CORE__88 = _
            " ALTER TABLE tb_siti_tabelle DROP COLUMN " + _
			"	tab_field_titolo, " + _
			"	tab_field_descrizione, " + _
			"	tab_field_url, " + _
			"	tab_field_meta_keywords, " + _
			"	tab_field_meta_description, " + _
			"	tab_field_alt " + _
			" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 89
'...........................................................................................
'corregge dimensione campo url su indice per sql server
'aggiunge id del sito a cui corrisponde il nodo dell'indice
'aggiunge struttura di gestione degli url rediretti
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__89(conn)
	Aggiornamento__FRAMEWORK_CORE__89 = _
            " ALTER TABLE tb_contents_index ALTER COLUMN idx_link_url_rw_it	" + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NULL ; " + _
			" ALTER TABLE tb_contents_index ALTER COLUMN idx_link_url_rw_en	" + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NULL ; " + _
			" ALTER TABLE tb_contents_index ALTER COLUMN idx_link_url_rw_fr	" + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NULL ; " + _
			" ALTER TABLE tb_contents_index ALTER COLUMN idx_link_url_rw_de	" + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NULL ; " + _
			" ALTER TABLE tb_contents_index ALTER COLUMN idx_link_url_rw_es	" + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NULL ; " + _
			" ALTER TABLE tb_contents_index ADD " + _
			"	idx_webs_id INT NULL " + _
			" ; "
	
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__89 = Aggiornamento__FRAMEWORK_CORE__89 + _
			" UPDATE tb_contents_index i INNER JOIN tb_pagineSito p ON i.idx_link_pagina_id = p.id_pagineSito SET " + _
			"	idx_webs_id = id_Web" + _
			" ; "
	else
		Aggiornamento__FRAMEWORK_CORE__89 = Aggiornamento__FRAMEWORK_CORE__89 + _
			" UPDATE tb_contents_index SET " + _
			"	idx_webs_id = (SELECT id_Web FROM tb_pagineSito WHERE tb_pagineSito.id_pagineSito = tb_contents_index.idx_link_pagina_id)" & _
			" 	WHERE " + SQL_IfIsNull(conn, "idx_link_pagina_id", "0") + ">0 " + _
			" ; "
	end if
	
	Aggiornamento__FRAMEWORK_CORE__89 = Aggiornamento__FRAMEWORK_CORE__89 + _
			AddInsModRelations(conn, "tb_contents_index", "idx") + _
			SQL_AddForeignKey(conn, "tb_contents_index", "idx_webs_id", "tb_webs", "id_webs", false, "") + _
			" CREATE TABLE " + SQL_Dbo(Conn) + "rel_index_url_redirect ( " + _
            "   riu_id " + SQL_PrimaryKey(conn, "rel_index_url_redirect") + ", " + _
			"	riu_idx_id INT NOT NULL, " + _
			"	riu_url " + SQL_CharField(Conn, IIF(DB_Type(conn) = DB_SQL, 500, 255)) + " NOT NULL, " + _
			"	riu_lingua " + IIF(DB_Type(conn) = DB_SQL, " varchar", " TEXT") + "(2) NOT NULL, " + _
			AddInsModFields("riu") + _
			" ); " + _
			AddInsModRelations(conn, "rel_index_url_redirect", "riu") + _
			SQL_AddForeignKey(conn, "rel_index_url_redirect", "riu_idx_id", "tb_contents_index", "idx_id", true, "") + _
			SQL_AddForeignKey(conn, "rel_index_url_redirect", "riu_lingua", "tb_cnt_lingue", "lingua_codice", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 90
'...........................................................................................
'rigenera url rewrited salvando il sito
'...........................................................................................
function AggiornamentoSpeciale__FRAMEWORK_CORE__90(DB, rs, version)
	dim sql
	'esegue un aggiornamento fasullo per aumentare la versione
	sql = "SELECT * FROM AA_Versione"
	CALL DB.Execute(sql, version)
	
	if DB.last_update_executed then
		CALL DB.RebuildIndex_OperazioniRicorsive()
	end if
	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 91
'...........................................................................................
'aggiunge campo di attivazione url rewriting
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__91(conn)
	Aggiornamento__FRAMEWORK_CORE__91 = _
            " ALTER TABLE tb_webs " & SQL_AddColumn(conn) & " URL_rewriting_attivo BIT NULL ; " + _
			" UPDATE tb_webs SET URL_rewriting_attivo = 0 ; " + _
			" ALTER TABLE tb_webs ALTER COLUMN URL_rewriting_attivo BIT NOT NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 92
'...........................................................................................
'imposta flag "url principale" per le voci dell'indice esistenti
'diventano tutti principali tranne: raggruppamenti e voci figlie di raggruppamenti, ma senza figli
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__92(conn)
	Aggiornamento__FRAMEWORK_CORE__92 = _
			" UPDATE tb_contents_index SET idx_principale = 1 ; " + _
		    " UPDATE tb_contents_index SET idx_principale = 0 " + _
			" 	WHERE idx_content_id IN (SELECT co_id FROM tb_contents WHERE co_F_table_id IN (SELECT tab_id FROM tb_siti_Tabelle WHERE tab_name LIKE 'tb_contents_index')) ; "
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__92 = Aggiornamento__FRAMEWORK_CORE__92 + _
			" UPDATE (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + _
			       " INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + _
			" SET idx_principale = 0 " + _
			" WHERE (SELECT COUNT (*) FROM v_indice WHERE NOT idx_principale AND  ( ',' +  tb_contents_index.idx_tipologie_padre_lista  +  ',' LIKE '%,' & idx_id & ',%' ) ) > 0 " + _
			"  AND (SELECT COUNT(*) FROM v_indice v_idx_duplicati WHERE v_idx_duplicati.co_f_key_id = tb_contents.co_f_key_id  AND v_idx_duplicati.tab_name LIKE tb_siti_tabelle.tab_name)>1"
	else
		Aggiornamento__FRAMEWORK_CORE__92 = Aggiornamento__FRAMEWORK_CORE__92 + _
			" UPDATE tb_contents_index SET idx_principale = 0 " + _
			" FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + _
			"	   INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + _
			" WHERE ( SELECT COUNT (*) FROM v_indice WHERE NOT  ISNULL(idx_principale,0) = 1  AND  ( ',' +  tb_contents_index.idx_tipologie_padre_lista  +  ',' LIKE '%,' + CAST(idx_id AS NVARCHAR) + ',%' ) ) > 0 " + _
			" 	   AND (SELECT COUNT(*) FROM v_indice v_idx_duplicati WHERE v_idx_duplicati.co_f_key_id = tb_contents.co_f_key_id  AND v_idx_duplicati.tab_name LIKE tb_siti_tabelle.tab_name)>1 "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 93
'...........................................................................................
'modifica vista per generazione sitemap
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__93(conn)
	Aggiornamento__FRAMEWORK_CORE__93 = _
		DropObject(conn, "v_indice_sitemap", "VIEW") + _
		" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
		"    SELECT DISTINCT " + vbCrLf + _
		"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND " & SQL_IfIsNull(conn, "idx_link_url_rw_it", "''") & "<>'' " , _
							"URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_it " + vbCrLf, _
							"URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_IT " + vbCrLf ) + ") AS uri, " + vbCrLF + _
		"        idx_modData, id_webs, idx_id, idx_livello " + vbCrLf + _
		"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
		"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
		"    WHERE (idx_link_url_IT <> '') "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 94
'...........................................................................................
'modifica vista per generazione sitemap
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__94(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__94 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__94 = Aggiornamento__FRAMEWORK_CORE__94 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND " & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__94 = Aggiornamento__FRAMEWORK_CORE__94 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE idx_principale OR NOT idx_foglia "
	else
		Aggiornamento__FRAMEWORK_CORE__94 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__94 = Aggiornamento__FRAMEWORK_CORE__94 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND " & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") "
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 95
'...........................................................................................
'corregge descrizione contenuti del next-web nelle tabelle per indice.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__95(conn)
	Aggiornamento__FRAMEWORK_CORE__95 = _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_colore = '#000000', " + _
		"	tab_from_sql = 'tb_webs', " + _
		"	tab_field_titolo_it = 'titolo_it', " + _
		"	tab_field_titolo_en = 'titolo_en', " + _
		"	tab_field_titolo_fr = 'titolo_fr', " + _
		"	tab_field_titolo_de = 'titolo_de', " + _
		"	tab_field_titolo_es = 'titolo_es', " + _
		"	tab_field_titolo_alt_it	 = '', " + _
		"	tab_field_titolo_alt_en	 = '', " + _
		"	tab_field_titolo_alt_fr	 = '', " + _
		"	tab_field_titolo_alt_de	 = '', " + _
		"	tab_field_titolo_alt_es	 = '', " + _
		"	tab_field_descrizione_it = 'META_description_it', " + _
		"	tab_field_descrizione_en = 'META_description_en', " + _
		"	tab_field_descrizione_fr = 'META_description_fr', " + _
		"	tab_field_descrizione_de = 'META_description_de', " + _
		"	tab_field_descrizione_es = 'META_description_es', " + _
		"	tab_field_meta_keywords_it = 'META_keywords_it', " + _
		"	tab_field_meta_keywords_en = 'META_keywords_en', " + _
		"	tab_field_meta_keywords_fr = 'META_keywords_fr', " + _
		"	tab_field_meta_keywords_de = 'META_keywords_de', " + _
		"	tab_field_meta_keywords_es = 'META_keywords_es', " + _
		"	tab_field_meta_description_it = 'META_description_it', " + _
		"	tab_field_meta_description_en = 'META_description_en', " + _
		"	tab_field_meta_description_fr = 'META_description_fr', " + _
		"	tab_field_meta_description_de = 'META_description_de', " + _
		"	tab_field_meta_description_es = 'META_description_es' " + _
		" WHERE tab_name LIKE 'tb_webs' ; " + _
		_
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_colore = '#FF0000', " + _
		"	tab_from_sql = 'tb_pagineSito', " + _
		"	tab_field_titolo_it = 'nome_ps_it', " + _
		"	tab_field_titolo_en = 'nome_ps_en', " + _
		"	tab_field_titolo_fr = 'nome_ps_fr', " + _
		"	tab_field_titolo_de = 'nome_ps_de', " + _
		"	tab_field_titolo_es = 'nome_ps_es', " + _
		"	tab_field_titolo_alt_it	 = '', " + _
		"	tab_field_titolo_alt_en	 = '', " + _
		"	tab_field_titolo_alt_fr	 = '', " + _
		"	tab_field_titolo_alt_de	 = '', " + _
		"	tab_field_titolo_alt_es	 = '', " + _
		"	tab_field_descrizione_it = 'PAGE_description_it', " + _
		"	tab_field_descrizione_en = 'PAGE_description_en', " + _
		"	tab_field_descrizione_fr = 'PAGE_description_fr', " + _
		"	tab_field_descrizione_de = 'PAGE_description_de', " + _
		"	tab_field_descrizione_es = 'PAGE_description_es', " + _
		"	tab_field_meta_keywords_it = 'PAGE_keywords_it', " + _
		"	tab_field_meta_keywords_en = 'PAGE_keywords_en', " + _
		"	tab_field_meta_keywords_fr = 'PAGE_keywords_fr', " + _
		"	tab_field_meta_keywords_de = 'PAGE_keywords_de', " + _
		"	tab_field_meta_keywords_es = 'PAGE_keywords_es', " + _
		"	tab_field_meta_description_it = 'PAGE_description_it', " + _
		"	tab_field_meta_description_en = 'PAGE_description_en', " + _
		"	tab_field_meta_description_fr = 'PAGE_description_fr', " + _
		"	tab_field_meta_description_de = 'PAGE_description_de', " + _
		"	tab_field_meta_description_es = 'PAGE_description_es' " + _
		" WHERE tab_name LIKE 'tb_pagineSito' ; " + _
		_
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_colore = '#808080', " + _
		"	tab_from_sql = 'v_indice', " + _
		"	tab_field_foto_thumb = 'co_foto_thumb', " + _
		"	tab_field_foto_zoom = 'co_foto_zoom', " + _
		"	tab_field_visibile = 'co_visibile', " + _
		"	tab_field_ordine = 'co_ordine', " + _
		"	tab_field_data_pubblicazione = 'co_data_pubblicazione', " + _
		"	tab_field_data_scadenza = 'co_data_scadenza', " + _
		"	tab_parametro = '', " + _
		"	tab_field_url_it = '', " + _
		"	tab_field_url_en = '', " + _
		"	tab_field_url_fr = '', " + _
		"	tab_field_url_de = '', " + _
		"	tab_field_url_es = '', " + _
		"	tab_field_titolo_it = 'co_titolo_it', " + _
		"	tab_field_titolo_en = 'co_titolo_en', " + _
		"	tab_field_titolo_fr = 'co_titolo_fr', " + _
		"	tab_field_titolo_de = 'co_titolo_de', " + _
		"	tab_field_titolo_es = 'co_titolo_es', " + _
		"	tab_field_titolo_alt_it	 = 'co_alt_it', " + _
		"	tab_field_titolo_alt_en	 = 'co_alt_en', " + _
		"	tab_field_titolo_alt_fr	 = 'co_alt_fr', " + _
		"	tab_field_titolo_alt_de	 = 'co_alt_de', " + _
		"	tab_field_titolo_alt_es	 = 'co_alt_es', " + _
		"	tab_field_descrizione_it = 'co_descrizione_it', " + _
		"	tab_field_descrizione_en = 'co_descrizione_en', " + _
		"	tab_field_descrizione_fr = 'co_descrizione_fr', " + _
		"	tab_field_descrizione_de = 'co_descrizione_de', " + _
		"	tab_field_descrizione_es = 'co_descrizione_es', " + _
		"	tab_field_meta_keywords_it = 'co_meta_keywords_it', " + _
		"	tab_field_meta_keywords_en = 'co_meta_keywords_en', " + _
		"	tab_field_meta_keywords_fr = 'co_meta_keywords_fr', " + _
		"	tab_field_meta_keywords_de = 'co_meta_keywords_de', " + _
		"	tab_field_meta_keywords_es = 'co_meta_keywords_es', " + _
		"	tab_field_meta_description_it = 'co_meta_description_it', " + _
		"	tab_field_meta_description_en = 'co_meta_description_en', " + _
		"	tab_field_meta_description_fr = 'co_meta_description_fr', " + _
		"	tab_field_meta_description_de = 'co_meta_description_de', " + _
		"	tab_field_meta_description_es = 'co_meta_description_es' " + _
		" WHERE tab_name LIKE '" & tabRaggruppamentoTable & "' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 96
'...........................................................................................
'esegue aggiornamento dati per correzione dati dell'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__96(conn)
	if DB_Type(conn) = DB_ACCESS then
		 Aggiornamento__FRAMEWORK_CORE__96 = _
			" UPDATE tb_contents_index i INNER JOIN tb_pagineSito p ON i.idx_link_pagina_id = p.id_pagineSito SET " + _
			"	idx_webs_id = id_Web" + _
			" ; " + _
			" UPDATE tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id " + _
			" 	SET idx_webs_id = co_F_key_id " + _
			" WHERE co_F_table_id = " & index.GetTable(tabSitoTable)
	else
		 Aggiornamento__FRAMEWORK_CORE__96 = _
			" UPDATE tb_contents_index SET " + _
			"	idx_webs_id = (SELECT id_Web FROM tb_pagineSito WHERE tb_pagineSito.id_pagineSito = tb_contents_index.idx_link_pagina_id)" & _
			" 	WHERE " + SQL_IfIsNull(conn, "idx_link_pagina_id", "0") + ">0 " + _
			" ; " + _
			" UPDATE tb_contents_index SET " + _
			"	idx_webs_id = (SELECT co_F_key_id FROM tb_contents WHERE co_id = tb_contents_index.idx_content_id ) " + _
			"   WHERE idx_content_id IN (SELECT co_id FROM tb_contents WHERE co_F_table_ID=" & index.GetTable(tabSitoTable) & " )"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 97
'...........................................................................................
'corregge descrizione tb_web per l'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__97(conn)
	Aggiornamento__FRAMEWORK_CORE__97 = _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_colore = '#000000', " + _
		"	tab_from_sql = 'tb_webs', "
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__97 = Aggiornamento__FRAMEWORK_CORE__97 + _
			"	tab_field_url_it = 'id_home_page' "
	else
		Aggiornamento__FRAMEWORK_CORE__97 = Aggiornamento__FRAMEWORK_CORE__97 + _
			"	tab_field_url_it = '(CASE WHEN sito_in_aggiornamento=1 THEN sito_in_aggiornamento_pagina WHEN sito_in_costruzione=1 THEN sito_in_costruzione_pagina ELSE id_home_page END)'"
	end if
	Aggiornamento__FRAMEWORK_CORE__97 = Aggiornamento__FRAMEWORK_CORE__97 + _
		" WHERE tab_name LIKE 'tb_webs' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 98
'...........................................................................................
'02/07/2008
'...........................................................................................
'corregge vista per generazione sitemap per esclusione pagine protette
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__98(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__98 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__98 = Aggiornamento__FRAMEWORK_CORE__98 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND " & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__98 = Aggiornamento__FRAMEWORK_CORE__98 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata "
	else
		Aggiornamento__FRAMEWORK_CORE__98 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__98 = Aggiornamento__FRAMEWORK_CORE__98 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND " & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata")
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 99
'...........................................................................................
'corregge flag principale degli url dell'indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__99(conn)
	Aggiornamento__FRAMEWORK_CORE__99 = Aggiornamento__FRAMEWORK_CORE__92(conn)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 100
'...........................................................................................
'aggiunge gestione "sedi" ai contatti interni
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__100(conn)
	Aggiornamento__FRAMEWORK_CORE__100 = _
            " ALTER TABLE tb_Indirizzario " & SQL_AddColumn(conn) & " CntSede INT NULL ; " + _
			SQL_AddForeignKey(conn, "tb_Indirizzario", "CntSede", "tb_Indirizzario", "IdElencoIndirizzi", false, "CntIntSede") + _
			SQL_AddForeignKey(conn, "tb_Indirizzario", "CntRel", "tb_Indirizzario", "IdElencoIndirizzi", false, "CntIntPadre")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 101
'...........................................................................................
'imposta l'indice del sito a "non principale" per non fare confusione con l'home page che ha lo stesso url.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__101(conn)
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__101 = _
            " UPDATE (tb_contents_index i" + vbCrLf + _
			" INNER JOIN tb_contents c ON i.idx_content_id = c.co_id)" + vbCrLf + _
			" INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id" + vbCrLf + _
			" SET idx_principale = 0" + vbCrLf + _
			" WHERE tab_name LIKE 'tb_webs'"
	else
		Aggiornamento__FRAMEWORK_CORE__101 = _
            " UPDATE tb_contents_index SET idx_principale = 0" + vbCrLf + _
			" WHERE idx_content_id IN (" + vbCrLf + _
			"	SELECT co_id FROM tb_contents c INNER JOIN tb_siti_tabelle t ON c.co_F_table_id = t.tab_id" + vbCrLf + _
			"	WHERE tab_name LIKE 'tb_webs')"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 102
'...........................................................................................
'corregge visibilit&agrave; pagine su indice: visibili solo se con almeno un layer
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__102(conn)
	Aggiornamento__FRAMEWORK_CORE__102 = _
    	" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_visibile='" + SQL_If(conn, "(SELECT COUNT(*) FROM tb_layers INNER JOIN tb_pages ON tb_layers.id_pag = tb_pages.id_page WHERE id_PaginaSito = tb_paginesito.id_paginesito)>0", "1", "0") + "' " + _
		" WHERE tab_name LIKE 'tb_pagineSito' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 103
'...........................................................................................
'corregge vista per generazione sitemap per calcolo url home page
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__103(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__103 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__103 = Aggiornamento__FRAMEWORK_CORE__103 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__103 = Aggiornamento__FRAMEWORK_CORE__103 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_name <> 'tb_webs' "
	else
		Aggiornamento__FRAMEWORK_CORE__103 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__103 = Aggiornamento__FRAMEWORK_CORE__103 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"		   AND tab_name <> 'tb_webs'"
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 104
'...........................................................................................
'corregge obbligatorieta' campo su log_cnt_email
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__104(conn)
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__104 = _
			" ALTER TABLE log_cnt_email ADD COLUMN log_cnt_id_temp INT NULL ; " + _
			" UPDATE log_cnt_email SET log_cnt_id_temp = log_cnt_id ; " + _
			" ALTER TABLE log_cnt_email DROP COLUMN log_cnt_id ; " + _
			" ALTER TABLE log_cnt_email ADD COLUMN log_cnt_id INT NULL ; " + _
			" UPDATE log_cnt_email SET log_cnt_id = log_cnt_id_temp ; " + _
			" ALTER TABLE log_cnt_email DROP COLUMN log_cnt_id_temp ; "
	else
		Aggiornamento__FRAMEWORK_CORE__104 = _
    		" ALTER TABLE log_cnt_email ALTER COLUMN log_cnt_id INT NULL "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 105
'...........................................................................................
'corregge descrizione contenuti del next-web nelle tabelle per indice.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__105(conn)
	Aggiornamento__FRAMEWORK_CORE__105 = _
		" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_url_it = 'id_home_page' " + _
		" WHERE tab_name LIKE 'tb_webs' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 106
'...........................................................................................
'aggiunge la relazione tra i formati delle immagini del NextWeb5 e tb_siti_tabelle
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__106(conn)
	Aggiornamento__FRAMEWORK_CORE__106 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"rel_immaginiFormati (" + vbCrLf + _
		"	rif_id "& SQL_PrimaryKey(conn, "rel_immaginiFormati") + ", " + vbCrLf + _
		"	rif_imf_id INT NOT NULL, " + vbCrLf + _
		"	rif_tab_id INT NOT NULL);" + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_immaginiFormati", "rif_imf_id", "tb_immaginiFormati", "imf_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_immaginiFormati", "rif_tab_id", "tb_siti_tabelle", "tab_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 107
'...........................................................................................
'aggiunge i campi thumb e zoom a tb_siti_tabelle
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__107(conn)
	Aggiornamento__FRAMEWORK_CORE__107 = _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_siti_tabelle ADD" + vbCrLf + _
		"	tab_thumb INT NULL," + vbCrLf + _
		"	tab_zoom INT NULL;" + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_siti_tabelle", "tab_thumb", "tb_immaginiFormati", "imf_id", false, "thumb") + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_siti_tabelle", "tab_zoom", "tb_immaginiFormati", "imf_id", false, "zoom") + vbCrLf + _
		" UPDATE tb_siti_tabelle SET tab_thumb = 0, tab_zoom = 0"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 108
'...........................................................................................
'corregge la correzzione su visibilit&agrave; pagine su indice: visibili solo se con almeno un layer
'ripara alle rogne generate con l'aggiornamento 102
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__108(conn)
	Aggiornamento__FRAMEWORK_CORE__108 = _
    	" UPDATE tb_siti_tabelle SET " + _
		"	tab_field_visibile='' " + _
		" WHERE tab_name LIKE 'tb_pagineSito' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 109
'...........................................................................................
'aggiunge proprieta all'amministratore.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__109(conn)
	Aggiornamento__FRAMEWORK_CORE__109 = _
    	" ALTER TABLE "& SQL_Dbo(conn) &"tb_admin ADD" + vbCrLf + _
		" 	admin_dir " + SQL_CharField(Conn, 255) + "," + vbCrLf + _
		" 	admin_telefono " + SQL_CharField(Conn, 250) + "," + vbCrLf + _
		" 	admin_fax " + SQL_CharField(Conn, 250) + "," + vbCrLf + _
		" 	admin_cell " + SQL_CharField(Conn, 250) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 110
'...........................................................................................
'aggiunge tabella dei servizi di comunicazione.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__110(conn)
	Aggiornamento__FRAMEWORK_CORE__110 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_servizi (" + vbCrLf + _
		"	serv_id "& SQL_PrimaryKey(conn, "tb_servizi") + ", " + vbCrLf + _
		"	serv_nome "+ SQL_CharField(Conn, 50) + ");" + vbCrLf + _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_email ADD" + vbCrLf + _
		"	email_servizio_id INT NULL; " + vbCrLf + _
		"INSERT INTO tb_servizi (serv_nome) VALUES ('E-MAIL');" + vbCrLf + _
		"INSERT INTO tb_servizi (serv_nome) VALUES ('SMS');" + vbCrLf + _
		"INSERT INTO tb_servizi (serv_nome) VALUES ('FAX');" + vbCrLf + _
		"UPDATE tb_email SET email_servizio_id = 1;" + _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_email ALTER COLUMN email_servizio_id INT NOT NULL;" + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_email", "email_servizio_id", "tb_servizi", "serv_id", true, "") + ";" + vbCrLf 
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 111
'...........................................................................................
'aggiunge campo per formati immagini
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__111(conn)
	Aggiornamento__FRAMEWORK_CORE__111 = _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_immaginiFormati ADD" + vbCrLf + _
		"	imf_dimensioniMax BIT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 112
'...........................................................................................
'aggiunge parametri tipizzati agli applicativi del NextPassport
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__112(conn)
	Aggiornamento__FRAMEWORK_CORE__112 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_siti_descrittori_raggruppamenti ( " + vbCrLf + _
		"	sdr_id "& SQL_PrimaryKey(conn, "tb_siti_descrittori_raggruppamenti") + ", " + vbCrLf + _
		"	sdr_titolo_it " + SQL_CharField(Conn, 250) + " NULL , " + vbCrLF + _
		"	sdr_titolo_en " + SQL_CharField(Conn, 250) + " NULL , " + vbCRLF + _
		"	sdr_titolo_fr " + SQL_CharField(Conn, 250) + " NULL , " + vbCrLF + _
		"	sdr_titolo_es " + SQL_CharField(Conn, 250) + " NULL , " + vbCrLf + _
		"	sdr_titolo_de " + SQL_CharField(Conn, 250) + " NULL , " + vbCrLf + _
		"	sdr_ordine int NULL, " + vbCrLf + _
		"	sdr_personalizzato bit NULL" + vbCrLf + _
		" );" + vbCrLf + _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_siti_descrittori (" + vbCrLf + _
		"	sid_id "& SQL_PrimaryKey(conn, "tb_siti_descrittori") + ", " + vbCrLf + _
		"	sid_raggruppamento_id int NULL ," + vbCrLf + _
		"	sid_codice " + SQL_CharField(Conn, 50) + " NOT NULL ," + vbCrLf + _
		"	sid_nome_it " + SQL_CharField(Conn, 255) + " NULL ," + vbCrLf + _
		"	sid_nome_en " + SQL_CharField(Conn, 255) + " NULL ," + vbCrLf + _
		"	sid_nome_fr " + SQL_CharField(Conn, 255) + " NULL ," + vbCrLf + _
		"	sid_nome_es " + SQL_CharField(Conn, 255) + " NULL ," + vbCrLf + _
		"	sid_nome_de " + SQL_CharField(Conn, 255) + " NULL ," + vbCrLf + _
		"	sid_unita_it " + SQL_CharField(Conn, 50) + " NULL ," + vbCrLf + _
		"	sid_unita_en " + SQL_CharField(Conn, 50) + " NULL ," + vbCrLf + _
		"	sid_unita_fr " + SQL_CharField(Conn, 50) + " NULL ," + vbCrLf + _
		"	sid_unita_es " + SQL_CharField(Conn, 50) + " NULL ," + vbCrLf + _
		"	sid_unita_de " + SQL_CharField(Conn, 50) + " NULL ," + vbCrLf + _
		"	sid_tipo int NULL ," + vbCrLf + _
		"	sid_principale bit NULL ," + vbCrLf + _
		"	sid_img " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	sid_admin bit NULL," + vbCrLf + _
		"	sid_personalizzato bit NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		" CREATE TABLE "& SQL_Dbo(conn) &"rel_siti_descrittori (" + vbCrLf + _
		"	rsd_id "& SQL_PrimaryKey(conn, "rel_siti_descrittori") + ", " + vbCrLf + _
		"	rsd_sito_id int NULL ," + vbCrLf + _
		"	rsd_descrittore_id int NULL ," + vbCrLf + _
		"	rsd_valore_it " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	rsd_valore_en " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	rsd_valore_fr " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	rsd_valore_es " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	rsd_valore_de " + SQL_CharField(Conn, 250) + " NULL ," + vbCrLf + _
		"	rsd_memo_it " + SQL_CharField(Conn, 0) + " NULL ," + vbCrLf + _
		"	rsd_memo_en " + SQL_CharField(Conn, 0) + " NULL ," + vbCrLf + _
		"	rsd_memo_fr " + SQL_CharField(Conn, 0) + " NULL ," + vbCrLf + _
		"	rsd_memo_es " + SQL_CharField(Conn, 0) + " NULL ," + vbCrLf + _
		"	rsd_memo_de " + SQL_CharField(Conn, 0) + " NULL " + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_siti_descrittori", "sid_raggruppamento_id", "tb_siti_descrittori_raggruppamenti", "sdr_id", false, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_descrittori", "rsd_descrittore_id", "tb_siti_descrittori", "sid_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 113
'...........................................................................................
'aggiunge parametri tipizzati agli applicativi del NextPassport
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__113(conn)
	Aggiornamento__FRAMEWORK_CORE__113 = _
		SQL_AddForeignKey(conn, "rel_siti_descrittori", "rsd_sito_id", "tb_siti", "id_sito", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 114
'...........................................................................................
'converte i parametri del NextGallery e del NextPassport
'...........................................................................................
Sub AggiornamentoSpeciale__FRAMEWORK_CORE__114(DB, rs, version)
	CALL DB.Execute("SELECT * FROM aa_versione", version)
	if DB.last_update_executed then
		dim siti
		siti = "1"
		if CIntero(GetValueList(DB.objconn, rs, "SELECT COUNT(*) FROM tb_siti WHERE id_sito = 27")) > 0 then
			siti = siti &",27"
		end if
		CALL ParametersImport(DB.objconn, siti)
	end if
End Sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 115
'...........................................................................................
'modifica vista v_indice, aggiungendo colonna per calcolo visibilita' dell'elemento
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__115(conn)
	Aggiornamento__FRAMEWORK_CORE__115 = _
		DropObject(conn, "v_indice", "VIEW") + _
		" CREATE VIEW " & SQL_Dbo(Conn) & "v_indice AS " + vbCrLf + _
		"    SELECT *, " + vbCrLF + _
		"    (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 			  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
							  "(" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
							  "("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto " + vbCrLF + _
		"    FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 116
'...........................................................................................
'aggiunge tabelle per gestione alert e notifiche
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__116(conn)
	Aggiornamento__FRAMEWORK_CORE__116 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_siti_eventi (" + vbCrLf + _
		"	sev_id "& SQL_PrimaryKey(conn, "tb_siti_eventi") + ", " + vbCrLf + _
		"	sev_sito_id INT NULL," + vbCrLf + _
		SQL_MultiLanguageField("sev_nome_<lingua> " + SQL_CharField(Conn, 250) + " NULL") +","+ vbCrLf + _
		"sev_codice " + SQL_CharField(Conn, 50) + " NULL" +","+ vbCrLf + _
		"	sev_abilitato BIT NULL," + vbCrLf + _
		"	sev_multisito BIT NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_siti_eventi", "sev_sito_id", "tb_siti", "id_sito", true, "") + vbCrLf + _
		" CREATE TABLE "& SQL_Dbo(conn) &"rel_siti_eventi (" + vbCrLf + _
		"	rse_id "& SQL_PrimaryKey(conn, "rel_siti_eventi") + ", " + vbCrLf + _
		"	rse_evento_id INT NULL," + vbCrLf + _
		"	rse_web_id INT NULL," + vbCrLf + _
		"	rse_email_abilitato BIT NULL," + vbCrLf + _
		"	rse_email_admin_id INT NULL," + vbCrLf + _
		"	rse_email_admin_invio BIT NULL," + vbCrLf + _
		"	rse_email_utenti_invio BIT NULL," + vbCrLf + _
		SQL_MultiLanguageField("rse_email_oggetto_<lingua> " + SQL_CharField(Conn, 250) + " NULL") +","+ vbCrLf + _
		SQL_MultiLanguageField("rse_email_testo_<lingua> " + SQL_CharField(Conn, 0) + " NULL") +","+ vbCrLf + _
		"	rse_email_paginaId INT NULL," + vbCrLf + _
		"	rse_sms_abilitato BIT NULL," + vbCrLf + _
		"	rse_sms_admin_id INT NULL," + vbCrLf + _
		"	rse_sms_admin_invio BIT NULL," + vbCrLf + _
		"	rse_sms_utenti_invio BIT NULL," + vbCrLf + _
		SQL_MultiLanguageField("rse_sms_testo_<lingua> " + SQL_CharField(Conn, 160) + " NULL") +","+ vbCrLf + _
		"	rse_fax_abilitato BIT NULL," + vbCrLf + _
		"	rse_fax_admin_id INT NULL," + vbCrLf + _
		"	rse_fax_admin_invio BIT NULL," + vbCrLf + _
		"	rse_fax_utenti_invio BIT NULL," + vbCrLf + _
		SQL_MultiLanguageField("rse_fax_oggetto_<lingua> " + SQL_CharField(Conn, 250) + " NULL") +","+ vbCrLf + _
		SQL_MultiLanguageField("rse_fax_testo_<lingua> " + SQL_CharField(Conn, 0) + " NULL") +","+ vbCrLf + _
		"	rse_fax_paginaId INT NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi", "rse_evento_id", "tb_siti_eventi", "sev_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi", "rse_web_id", "tb_webs", "id_webs", false, "") + vbCrLf + _
		" CREATE TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_admin (" + vbCrLf + _
		"	rea_id "& SQL_PrimaryKey(conn, "rel_siti_eventi_destinatari") + ", " + vbCrLf + _
		"	rea_sitoevento_id INT NULL," + vbCrLf + _
		"	rea_servizio_id INT NULL," + vbCrLf + _
		"	rea_admin_id INT NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_admin", "rea_sitoevento_id", "rel_siti_eventi", "rse_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_admin", "rea_admin_id", "tb_admin", "id_admin", true, "") + vbCrLf + _
		" CREATE TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_contatti (" + vbCrLf + _
		"	rec_id "& SQL_PrimaryKey(conn, "rel_siti_eventi_contatti") + ", " + vbCrLf + _
		"	rec_sitoevento_id INT NULL," + vbCrLf + _
		"	rec_servizio_id INT NULL," + vbCrLf + _
		"	rec_contatto_id INT NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_contatti", "rec_sitoevento_id", "rel_siti_eventi", "rse_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_contatti", "rec_contatto_id", "tb_indirizzario", "idElencoIndirizzi", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 116 bis
'...........................................................................................
'aggiunge relazioni con tabella servizi
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__116_bis(conn)
		Aggiornamento__FRAMEWORK_CORE__116_bis = _
			SQL_AddForeignKey(conn, "rel_siti_eventi_admin", "rea_servizio_id", "tb_servizi", "serv_id", true, "") + vbCrLf + _
			SQL_AddForeignKey(conn, "rel_siti_eventi_contatti", "rec_servizio_id", "tb_servizi", "serv_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 117
'...........................................................................................
' Annulla le relazioni introdotte da tb_servizi.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__117(conn)
		Aggiornamento__FRAMEWORK_CORE__117 = _
		SQL_RemoveForeignKey(conn, "rel_siti_eventi_admin", "rea_servizio_id", "tb_servizi",  true, "") + vbCrLf + _
		SQL_RemoveForeignKey(conn, "rel_siti_eventi_contatti", "rec_servizio_id", "tb_servizi",  true, "") + vbCrLf + _
		SQL_RemoveForeignKey(conn, "tb_email", "email_servizio_id", "tb_servizi", true, "") + ";" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 118
'...........................................................................................
'aggiunge tabella dei servizi di comunicazione.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__118(conn)
		Aggiornamento__FRAMEWORK_CORE__118 = _
		"ALTER TABLE "& SQL_Dbo(conn) &"tb_email DROP COLUMN email_servizio_id;" + vbCrLf + _
		DropObject(conn, "tb_servizi", "TABLE")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 119
'...........................................................................................
'aggiunge tabella dei servizi di comunicazione.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__119(conn)
	Aggiornamento__FRAMEWORK_CORE__119 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_tipi_messaggi (" + vbCrLf + _
		"	tipi_messaggi_id "& SQL_PrimaryKey(conn, "tb_tipi_messaggi") + ", " + vbCrLf + _
		"	tipi_messaggi_nome "+ SQL_CharField(Conn, 50) + ");" + vbCrLf + _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_email ADD" + vbCrLf + _
		"	email_tipi_messaggi_id INT NULL; " + vbCrLf + _
		" ALTER TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_admin ADD" + vbCrLf + _
		"	rea_tipo_messaggio_id INT NULL; " + vbCrLf + _
		" UPDATE rel_siti_eventi_admin SET rea_tipo_messaggio_id = rea_servizio_id ; " + _
		" ALTER TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_admin DROP COLUMN rea_servizio_id; " + vbCrLf + _
		" ALTER TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_contatti ADD" + vbCrLf + _
		"	rec_tipo_messaggio_id INT NULL; " + vbCrLf + _
		" UPDATE rel_siti_eventi_contatti SET rec_tipo_messaggio_id = rec_servizio_id ; " + _
		" ALTER TABLE "& SQL_Dbo(conn) &"rel_siti_eventi_contatti DROP COLUMN rec_servizio_id; " + vbCrLf + _
		"INSERT INTO tb_tipi_messaggi (tipi_messaggi_nome) VALUES ('E-MAIL');" + vbCrLf + _
		"INSERT INTO tb_tipi_messaggi (tipi_messaggi_nome) VALUES ('SMS');" + vbCrLf + _
		"INSERT INTO tb_tipi_messaggi (tipi_messaggi_nome) VALUES ('FAX');" + vbCrLf + _
		"UPDATE tb_email SET email_tipi_messaggi_id = 1;" + _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_email ALTER COLUMN email_tipi_messaggi_id INT NOT NULL;" + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_email", "email_tipi_messaggi_id", "tb_tipi_messaggi", "tipi_messaggi_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_admin", "rea_tipo_messaggio_id", "tb_tipi_messaggi", "tipi_messaggi_id", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "rel_siti_eventi_contatti", "rec_tipo_messaggio_id", "tb_tipi_messaggi", "tipi_messaggi_id", true, "") + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 120
'...........................................................................................
'aggiunge campo data modifica dei parametri
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__120(conn)
	Aggiornamento__FRAMEWORK_CORE__120 = _
		" ALTER TABLE tb_webs ADD" + vbCrLf + _
		"	webs_modData_parametri smalldatetime NULL;" + vbCrLf + _
		" UPDATE tb_webs SET webs_modData_parametri = " + SQL_Now(conn)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 121
'...........................................................................................
'aggiorna gli URL dell'indice (modifica gli url RW aggiungendo il prefisso della lingua)
'...........................................................................................
Sub AggiornamentoSpeciale__FRAMEWORK_CORE__121(DB, rs, version)
	CALL DB.Execute("SELECT * FROM aa_versione", version)
	if DB.last_update_executed then
		dim sql
		set index.conn = DB.objConn
		set index.content.conn = DB.objConn
		
		sql = " SELECT idx_id FROM tb_contents_index"& _
			  " WHERE idx_livello = 0"
		rs.open sql, DB.objConn, adOpenStatic, adLockOptimistic
		while not rs.eof
			index.operazioni_ricorsive_tipologia(rs("idx_id"))
			rs.movenext
		wend
		rs.close
	end if
End Sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 122
'...........................................................................................
'aggiunge campo data, ora e luogo alle news
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__122(conn)
	Aggiornamento__FRAMEWORK_CORE__122 = _
		" ALTER TABLE tb_news ADD" + vbCrLf + _
		"	news_agenda_data smalldatetime NULL, " + vbCrLf + _
		"	news_agenda_luogo " + SQL_CharField(Conn, 250) + " NULL, " + vbCrLf + _
		" 	news_agenda_sms_avviso BIT NULL, " + _
		"	news_agenda_sms_testo " + SQL_CharField(Conn, 160) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 123
'...........................................................................................
'	Nicola, 15/04/2009
'...........................................................................................
'aggiunge campo data di fine dell'evento alle news
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__123(conn)
	Aggiornamento__FRAMEWORK_CORE__123 = _
		" ALTER TABLE tb_news ADD" + vbCrLf + _
		"	news_agenda_ora smalldatetime NULL, " + _
		"	news_agenda_data_fine smalldatetime NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 124
'...........................................................................................
'	Nicola, 19/05/2009
'...........................................................................................
'aggiunge indicazione se il tagging e' abilitato per il contenuto
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__124(conn)
	Aggiornamento__FRAMEWORK_CORE__124 = _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	tab_tags_abilitati BIT NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 125
'...........................................................................................
'	Nicola, 19/05/2009
'...........................................................................................
'aggiunge vista per la gestione dei tags
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__125(conn)
	Aggiornamento__FRAMEWORK_CORE__125 = _
		" CREATE VIEW " & SQL_Dbo(Conn) & "v_tags AS " + vbCrLf + _
		"	SELECT * " + vbCrLf + _
		"		FROM (tb_contents_tags " + vbCrLf + _
		"		INNER JOIN rel_contents_tags ON tb_contents_tags.tag_id = rel_contents_tags.rct_tag_id ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 126
'...........................................................................................
'	Nicola, 20/05/2009
'...........................................................................................
'modifica struttura di base della tabella dei tag
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__126(conn)
	Aggiornamento__FRAMEWORK_CORE__126 = _
		" ALTER TABLE tb_contents_tags ADD " + _
		"	tag_value " + SQL_CharField(Conn, 250) + " NULL, " + _
		"	tag_lingua " + replace(SQL_CharField(Conn, 2), "nvarchar", "varchar") + " NULL " + _
		" ; " + _
		" ALTER TABLE tb_contents_tags DROP COLUMN tag_it ; " + _
		" ALTER TABLE tb_contents_tags DROP COLUMN tag_en ; " + _
		" ALTER TABLE tb_contents_tags DROP COLUMN tag_fr ; " + _
		" ALTER TABLE tb_contents_tags DROP COLUMN tag_de ; " + _
		" ALTER TABLE tb_contents_tags DROP COLUMN tag_es ; " + _
		SQL_AddForeignKey(conn, "tb_contents_tags", "tag_lingua", "tb_cnt_lingue", "lingua_codice", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 127
'...........................................................................................
'	Andrea, 20/05/2009
'...........................................................................................
'modifica struttura di base della tabella dei tag
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__127(conn)
	Aggiornamento__FRAMEWORK_CORE__127 = _
		" DELETE FROM tb_css_styles WHERE tb_css_styles.style_pseudoclass like '%:visited%';"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 128
'...........................................................................................
'	Simone, 16/07/2009
'...........................................................................................
' aggiunge log motore di ricerca
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__128(conn)
	Aggiornamento__FRAMEWORK_CORE__128 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"log_ricerche (" + vbCrLf + _
		"	lor_id "& SQL_PrimaryKey(conn, "log_ricerche") + ", " + vbCrLf + _
		"	lor_web_id INT NOT NULL, " + vbCrLf + _
		"	lor_index_id INT NULL, " + vbCrLf + _
		"	lor_object_id INT NULL, " + vbCrLf + _
		"	lor_ricerca "+ SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	lor_risultati_numero INT NULL, " + vbCrLf + _
		"	lor_data SMALLDATETIME NULL, " + vbCrLf + _
		"	lor_ip "+ SQL_CharField(Conn, 15) + "," + vbCrLf + _
		"	lor_request "+ SQL_CharField(Conn, 0) + ");" + vbCrLf + _
		SQL_AddForeignKey(conn, "log_ricerche", "lor_web_id", "tb_webs", "id_webs", true, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "log_ricerche", "lor_index_id", "tb_contents_index", "idx_id", false, "") + vbCrLf + _
		SQL_AddForeignKey(conn, "log_ricerche", "lor_object_id", "tb_objects", "id_objects", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 129
'...........................................................................................
'	Sergio, 20/05/2009
'...........................................................................................
'aggiunge tabella degli stati e la popola direttamente
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__129(conn)
	Dim fso, f, sqlbuffer,s,iso,nome, path
	path = server.mappath("stati.txt")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(path, 1, false)
	sqlbuffer = ""
    Do While f.AtEndOfStream <> True
      s = f.ReadLine
	  iso = left(s,3)
	  nome = right(s, len(s)-4)
	  sqlbuffer = sqlbuffer + " INSERT INTO "+ SQL_Dbo(conn) +"stati (codiceIso,nome) VALUES ('"+iso+"','"+replace(nome,"'","''")+"'); " +vbCrLf 
	Loop
	f.Close
	if DB_Type(conn) = DB_SQL then
	Aggiornamento__FRAMEWORK_CORE__129 = _
		" IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id(N'dbo.stati') AND type in (N'U')) " + _
		" CREATE TABLE "+ SQL_Dbo(conn) +"stati( " + vbCrLf + _
		"	codiceIso char (3) NOT NULL, " + vbCrLf + _
		"	nome nvarchar(255) NOT NULL); " + vbCrLf + _
		" " + sqlbuffer
	else
		Aggiornamento__FRAMEWORK_CORE__129 = _
		" CREATE TABLE "+ SQL_Dbo(conn) +"stati( " + vbCrLf+ _
		"	codiceIso char (3) NOT NULL, " + vbCrLf + _
		"	nome nvarchar(255) NOT NULL); " + vbCrLf + _
		" " + sqlbuffer
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 130
'...........................................................................................
'	Nicola, 23/07/2009
'...........................................................................................
'aggiunge campo partita_iva ai contatti
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__130(conn)
	Aggiornamento__FRAMEWORK_CORE__130 = _
		" ALTER TABLE tb_indirizzario ADD " + _
		"	partita_iva "+ SQL_CharField(Conn, 255)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 131
'...........................................................................................
'	Andrea, 04/08/2009
'...........................................................................................
'aggiunge campo descrizione e logo a FAQ_Categorie
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__131(conn)
	Aggiornamento__FRAMEWORK_CORE__131 = _
		" ALTER TABLE tb_FAQ_categorie ADD " + _
		SQL_MultiLanguageField(" cat_descr_<lingua> " + SQL_CharField(Conn, 0)) + ", " + _
		" cat_logo " + SQL_CharField(Conn, 255) + " NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 132
'...........................................................................................
'	Nicola, 05/08/2009
'...........................................................................................
' aggiunge tabella per la gestione dei meta tag aggiuntivi (ES: robotx, autenticazione motori di ricerca ecc..)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__132(conn)
	Aggiornamento__FRAMEWORK_CORE__132 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_webs_metatag (" + vbCrLf + _
		"	meta_id "& SQL_PrimaryKey(conn, "tb_webs_metatag") + ", " + vbCrLf + _
		"	meta_web_id INT NOT NULL, " + _ 
		"	meta_name "+ SQL_CharField(Conn, 255) + "," + vbCrLf + _
		"	meta_content "+ SQL_CharField(Conn, 255) + vbCrLf + _
		");" + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_webs_metatag", "meta_web_id", "tb_webs", "id_webs", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 133
'...........................................................................................
'	Giacomo, 20/08/2009
'...........................................................................................
'aggiunge campo a categorie delle news e dei link utili per indicazione ordine e pubblicazione
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__133(conn)
			Aggiornamento__FRAMEWORK_CORE__133 = _
				"ALTER TABLE tb_contents_index ADD " + _
				SQL_MultiLanguageField(" idx_titolo_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _
				SQL_MultiLanguageField(" idx_descrizione_<lingua> " + SQL_CharField(Conn, 0))
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 134
'...........................................................................................
'	Giacomo, 21/08/2009
'...........................................................................................
'aggiunge metatag
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__134(conn)
			Aggiornamento__FRAMEWORK_CORE__134 = _
				"INSERT INTO tb_webs_metatag(meta_web_id, meta_name, meta_content)" + vbCrLf + _
				"SELECT id_webs, 'robots', 'INDEX,FOLLOW'" + vbCrLf + _
				"FROM tb_webs"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 135
'...........................................................................................
'	Giacomo, 27/08/2009
'...........................................................................................
'aggiunge campi tab_ricercabile e tab_per_sitemap nella tabella tb_siti_tabelle
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__135(conn)
			Aggiornamento__FRAMEWORK_CORE__135 = _
				"ALTER TABLE tb_siti_tabelle ADD tab_ricercabile bit NULL, tab_per_sitemap bit NULL ;" + _
				"UPDATE tb_siti_tabelle SET tab_ricercabile=1, tab_per_sitemap=1 "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 136
'...........................................................................................
'	Simone, 30/09/2009
'...........................................................................................
'aggiunge campo per gestione modifica plugin.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__136(conn)
	Aggiornamento__FRAMEWORK_CORE__136 = _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_webs ADD" + vbCrLf + _
		"	webs_modData_plugin SMALLDATETIME NULL;" + vbCrLf + _
		" UPDATE tb_webs SET webs_modData_plugin = " + SQL_Now(conn)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 137
'...........................................................................................
'	Giacomo, 06/10/2009
'...........................................................................................
'   aggiunge campi id_logout_page_riservata e id_registrazione_page_riservata
'   nella tabella tb_webs.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__137(conn)
			Aggiornamento__FRAMEWORK_CORE__137 = _
				"ALTER TABLE tb_webs ADD id_logout_page_riservata INT NULL, " + vbCrLf + _
				" id_registrazione_page_riservata INT NULL;" + vbCrLf + _
				SQL_AddForeignKey(conn, "tb_webs", "id_logout_page_riservata", "tb_pagineSito", "id_pagineSito", false, "") + _
				SQL_AddForeignKey(conn, "tb_webs", "id_registrazione_page_riservata", "tb_pagineSito", "id_pagineSito", false, "2")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 138
'...........................................................................................
'	Nicola, Giacomo, 08/10/2009
'...........................................................................................
'   aggiunge funzione SQL per recuperare il nome completo di un elemento direttamente da una query.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__138(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__138 = _
			" CREATE FUNCTION dbo.fn_indice_nome_completo " + vbCrLf + _
			" ( " + vbCrLf + _
			" 	@idxId int, " + vbCrLf + _
			" 	@lingua nvarchar(2) " + vbCrLf + _
			" ) " + vbCrLf + _
			" RETURNS nvarchar(3000) " + vbCrLf + _
			" AS " + vbCrLf + _
			" BEGIN " + vbCrLf + _
			" 	DECLARE @idx_tipologie_padre_lista nvarchar(255) " + vbCrLf + _
			" 	SELECT @idx_tipologie_padre_lista = idx_tipologie_padre_lista FROM tb_contents_index WHERE idx_id = @IdxId " + vbCrLf + _
			" 	DECLARE RS CURSOR  " + vbCrLf + _
			" 	FOR	 " + vbCrLf + _
			" 		SELECT IsNull(co_titolo_it,''), IsNull(co_titolo_en,''), IsNull(co_titolo_fr,''), IsNull(co_titolo_de,''), IsNull(co_titolo_es,'') " + vbCrLf + _
			" 		FROM v_indice  " + vbCrLf + _
			" 		WHERE ',' + @idx_tipologie_padre_lista + ',' LIKE '%,' +  CAST(idx_id AS nvarchar(30)) + ',%' " + vbCrLf + _
			" 		ORDER BY idx_livello " + vbCrLf + _	
			" 	DECLARE @t_it nvarchar(255), @t_en nvarchar(255), @t_fr nvarchar(255), @t_de nvarchar(255), @t_es nvarchar(255)	 " + vbCrLf + _
			" 	DECLARE @full_it nvarchar(255), @full_en nvarchar(255), @full_fr nvarchar(255), @full_de nvarchar(255), @full_es nvarchar(255)	 " + vbCrLf + _
			" 	SELECT @full_it = '', @full_en = '', @full_fr = '', @full_de = '', @full_es = '' " + vbCrLf + _
			" 	OPEN RS " + vbCrLf + _
			" 	FETCH NEXT FROM RS INTO @t_it, @t_en, @t_fr, @t_de, @t_es " + vbCrLf + _
			" 	WHILE @@fetch_status <>-1 " + vbCrLf + _
			" 	BEGIN " + vbCrLf + _
			" 		if (@@fetch_status <>-2) " + vbCrLf + _
			" 		BEGIN " + vbCrLf + _
			" 			IF (IsNull(@t_it,'')<>'') " + vbCrLf + _
			" 				SET @full_it = RIGHT(@full_it + @t_it + ' - ', 3000) " + vbCrLf + _
			" 			IF (IsNull(@t_en,'')<>'') " + vbCrLf + _
			" 				SET @full_en = RIGHT(@full_en + @t_en + ' - ', 3000) " + vbCrLf + _
			" 			IF (IsNull(@t_fr,'')<>'') " + vbCrLf + _
			" 				SET @full_fr = RIGHT(@full_fr + @t_fr + ' - ', 3000) " + vbCrLf + _
			" 			IF (IsNull(@t_de,'')<>'') " + vbCrLf + _
			" 				SET @full_de = RIGHT(@full_de + @t_de + ' - ', 3000) " + vbCrLf + _
			" 			IF (IsNull(@t_es,'')<>'') " + vbCrLf + _
			" 				SET @full_es = RIGHT(@full_es + @t_es + ' - ', 3000) " + vbCrLf + _
					
			" 		END " + vbCrLf + _
			" 		FETCH NEXT FROM RS INTO @t_it, @t_en, @t_fr, @t_de, @t_es " + vbCrLf + _
			" 	END " + vbCrLf + _
			" 	CLOSE RS " + vbCrLf + _
			" 	DEALLOCATE RS " + vbCrLf + _
			" 	DECLARE @ret nvarchar(3000) " + vbCrLf + _
			" 	IF (LOWER(@lingua) = 'it' AND LEN(@full_it) > 2) " + vbCrLf + _
			" 		SET @ret = LEFT(@full_it, LEN(@full_it) - 2) " + vbCrLf + _
			" 	IF (LOWER(@lingua) = 'en' AND LEN(@full_en) > 2) " + vbCrLf + _
			" 		SET @ret = LEFT(@full_en, LEN(@full_en) - 2) " + vbCrLf + _
			" 	IF (LOWER(@lingua) = 'fr' AND LEN(@full_fr) > 2) " + vbCrLf + _
			" 		SET @ret = LEFT(@full_fr, LEN(@full_fr) - 2) " + vbCrLf + _
			" 	IF (LOWER(@lingua) = 'de' AND LEN(@full_de) > 2) " + vbCrLf + _
			" 		SET @ret = LEFT(@full_de, LEN(@full_de) - 2) " + vbCrLf + _
			" 	IF (LOWER(@lingua) = 'es' AND LEN(@full_es) > 2) " + vbCrLf + _
			" 		SET @ret = LEFT(@full_es, LEN(@full_es) - 2) " + vbCrLf + _	
			" 	RETURN  @ret " + vbCrLf + _
			" END "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 139
'...........................................................................................
'	Nicola, 08/10/2009
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap.
' 	imposta flag di visibilità su sitemap per tutti i contenuti.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__139(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__139 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__139 = Aggiornamento__FRAMEWORK_CORE__139 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__139 = Aggiornamento__FRAMEWORK_CORE__139 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND tab_name <> 'tb_webs' "
	else
		Aggiornamento__FRAMEWORK_CORE__139 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__139 = Aggiornamento__FRAMEWORK_CORE__139 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"		   AND tab_name <> 'tb_webs'"
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 140
'...........................................................................................
'	Nicola, Giacomo, 12/10/2009
'...........................................................................................
'   modifica funzione SQL per recuperare il nome completo di un elemento direttamente da una query.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__140(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__140 = _
			" ALTER FUNCTION dbo.fn_indice_nome_completo " + vbCrLf + _
			"  (  " + vbCrLf + _
			" 	  @idxId int,  " + vbCrLf + _
			" 	  @lingua nvarchar(2) " + vbCrLf + _
			"  )  " + vbCrLf + _
			" RETURNS nvarchar(3000) " + vbCrLf + _ 
			" AS  " + vbCrLf + _
			" BEGIN  " + vbCrLf + _
			" 	  DECLARE @currentIdxId int, @padreId int " + vbCrLf + _
			" 	  DECLARE @t_it nvarchar(255), @t_en nvarchar(255), @t_fr nvarchar(255), @t_de nvarchar(255), @t_es nvarchar(255)  " + vbCrLf + _
			" 	  DECLARE @full_it nvarchar(255), @full_en nvarchar(255), @full_fr nvarchar(255), @full_de nvarchar(255), @full_es nvarchar(255)  " + vbCrLf + _
			" 	  SELECT @full_it = '', @full_en = '', @full_fr = '', @full_de = '', @full_es = ''  " + vbCrLf + _
			" 	  SET @currentIdxId = @idxId " + vbCrLf + _
			" 	  WHILE @currentIdxId<>0 " + vbCrLf + _
			" 	  BEGIN " + vbCrLf + _
			" 			SELECT @currentIdxId = ISNULL(idx_id,0), " + vbCrLf + _
			" 					 @padreId = ISNULL(idx_padre_id,0), " + vbCrLf + _
			" 					 @t_it = IsNull(co_titolo_it,''), " + vbCrLf + _
			" 					 @t_en = IsNull(co_titolo_en,''), " + vbCrLf + _
			" 					 @t_fr = IsNull(co_titolo_fr,''), " + vbCrLf + _
			" 					 @t_de = IsNull(co_titolo_de,''), " + vbCrLf + _
			" 					 @t_es = IsNull(co_titolo_es,'') " + vbCrLf + _
			" 			FROM v_indice " + vbCrLf + _
			" 			WHERE idx_id = @currentIdxId " + vbCrLf + _
			" 			if (@currentIdxId <> 0) " + vbCrLf + _
			" 			BEGIN " + vbCrLf + _
			" 				  IF (IsNull(@t_it,'')<>'')  " + vbCrLf + _
			" 						SET @full_it = RIGHT(@t_it + ' - ' + @full_it, 3000)  " + vbCrLf + _
			" 				  IF (IsNull(@t_en,'')<>'')  " + vbCrLf + _
			" 						SET @full_en = RIGHT(@t_en + ' - ' + @full_en, 3000)  " + vbCrLf + _
			" 				  IF (IsNull(@t_fr,'')<>'')  " + vbCrLf + _
			" 						SET @full_fr = RIGHT(@t_fr + ' - ' + @full_fr, 3000)  " + vbCrLf + _
			" 				  IF (IsNull(@t_de,'')<>'')  " + vbCrLf + _
			" 						SET @full_de = RIGHT(@t_de + ' - ' + @full_de, 3000)  " + vbCrLf + _
			" 				  IF (IsNull(@t_es,'')<>'')  " + vbCrLf + _
			" 						SET @full_es = RIGHT(@t_es + ' - ' + @full_es, 3000) " + vbCrLf + _
			" 			END " + vbCrLf + _
			" 			SET @currentIdxId = @padreId " + vbCrLf + _
			" 	  END " + vbCrLf + _
			" 	  DECLARE @ret nvarchar(3000)  " + vbCrLf + _
			" 	  IF (LOWER(@lingua) = 'it' AND LEN(@full_it) > 2)  " + vbCrLf + _
			" 			SET @ret = LEFT(@full_it, LEN(@full_it) - 2)  " + vbCrLf + _
			" 	  IF (LOWER(@lingua) = 'en' AND LEN(@full_en) > 2)  " + vbCrLf + _
			" 			SET @ret = LEFT(@full_en, LEN(@full_en) - 2)  " + vbCrLf + _
			" 	  IF (LOWER(@lingua) = 'fr' AND LEN(@full_fr) > 2)  " + vbCrLf + _
			" 			SET @ret = LEFT(@full_fr, LEN(@full_fr) - 2)  " + vbCrLf + _
			" 	  IF (LOWER(@lingua) = 'de' AND LEN(@full_de) > 2)  " + vbCrLf + _
			" 			SET @ret = LEFT(@full_de, LEN(@full_de) - 2)  " + vbCrLf + _
			" 	  IF (LOWER(@lingua) = 'es' AND LEN(@full_es) > 2)  " + vbCrLf + _
			" 			SET @ret = LEFT(@full_es, LEN(@full_es) - 2)  " + vbCrLf + _
			" 	  RETURN  @ret  " + vbCrLf + _
			" END "
	end if
end function


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 141
'...........................................................................................
'	Nicola, 12/10/2009
'...........................................................................................
'   aggiunge campo per inserimento javascript su footer di pagina.
'	(lo aggiunge dopo codice attivazione google stats).
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__141(conn)
			Aggiornamento__FRAMEWORK_CORE__141 = _
				"ALTER TABLE tb_webs ADD " + _
				"	pagefooter_script " + SQL_CharField(Conn, 0) + ";" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 142
'...........................................................................................
'	Giacomo, 24/11/2009
'...........................................................................................
'   aggiunge un flag per sapere se il tag è stato generato automaticamente o no
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__142(conn)
			Aggiornamento__FRAMEWORK_CORE__142 = _
				"ALTER TABLE rel_contents_tags ADD " + _
				"	rct_autogenerato bit NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 143
'...........................................................................................
'	Giacomo, 30/11/2009
'...........................................................................................
'   
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__143(conn)
			Aggiornamento__FRAMEWORK_CORE__143 = _
					"	ALTER TABLE tb_siti_tabelle ADD " + _
					"	tab_priorita_base int NULL, " + _
					SQL_MultiLanguageField(" tab_field_return_url_<lingua> " + SQL_CharField(Conn, 500)) + ";" + _ 
					"	ALTER TABLE tb_contents_index ADD " + _
					"	idx_priorita int NULL; " + _
					"	ALTER TABLE tb_webs ADD " + _
					"	webs_modData_tabelle smalldatetime NULL; " + _
					"   UPDATE tb_webs SET webs_modData_tabelle = " + SQL_Now(conn) + " ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 144
'...........................................................................................
'	Giacomo, 01/12/2009
'...........................................................................................
'   aggiunge campo che conterrà l'elenco dei campi da utilizzare come tag 
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__144(conn)
			Aggiornamento__FRAMEWORK_CORE__144 = _
					"	ALTER TABLE tb_siti_tabelle ADD " + _
					SQL_MultiLanguageField(" tab_tags_fields_csv_<lingua> " + SQL_CharField(Conn, 500)) + "," + _
					SQL_MultiLanguageField(" tab_tags_fields_ssv_<lingua> " + SQL_CharField(Conn, 500)) + ";"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 145
'...........................................................................................
'	Giacomo, 15/02/2010
'...........................................................................................
'    
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__145(conn)
			Aggiornamento__FRAMEWORK_CORE__145 = _
					" ALTER TABLE tb_cnt_lingue ALTER COLUMN " + _
					" 		lingua_nome_IT " + SQL_CharField(Conn, 20) + ";" + _
					" ALTER TABLE tb_cnt_lingue ALTER COLUMN " + _
					" 		lingua_nome " + SQL_CharField(Conn, 20) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************

'AGGIORNAMENTO FRAMEWORK CORE 146
'...........................................................................................
'	Giacomo, 21/01/2010
'...........................................................................................
' SERIE DI FUNZIONI PER AGGIUNGERE I CAMPI PER UNA NUOVA LINGUA SU TUTTO IL FRAMEWORK
'...........................................................................................


function Update_language_NextCom_NextPassport(lingua_abbr)
	Update_language_NextCom_NextPassport = " ALTER TABLE rel_siti_descrittori ADD " + _
		  " 	rsd_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	rsd_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE rel_siti_eventi ADD " + _
		  " 	rse_email_oggetto_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	rse_email_testo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
		  " 	rse_sms_testo_" + lingua_abbr + " " + SQL_CharField(Conn, 160) + " NULL," + _
		  " 	rse_fax_oggetto_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	rse_fax_testo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE tb_siti_descrittori ADD " + _
		  " 	sid_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	sid_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 50) + " NULL;" + _
		  " ALTER TABLE tb_siti_descrittori_raggruppamenti ADD " + _
		  " 	sdr_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + _
		  " ALTER TABLE tb_siti_eventi ADD " + _
		  " 	sev_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + _
		  " ALTER TABLE tb_siti_tabelle ADD " + _
		  " 	tab_field_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_field_titolo_alt_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_field_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_field_codice_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_field_url_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_field_meta_keywords_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
		  " 	tab_field_meta_description_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
		  " 	tab_field_return_url_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_tags_fields_csv_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	tab_tags_fields_ssv_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;"
end function


function Update_language_NextWeb(lingua_abbr)
	Update_language_NextWeb = " ALTER TABLE tb_contents ADD " + _
							  " 	co_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	co_chiave_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	co_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	co_link_url_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	co_meta_keywords_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	co_meta_description_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	co_alt_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	co_link_url_rw_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
							  " ALTER TABLE tb_contents_index ADD " + _
							  " 	idx_link_url_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	idx_meta_keywords_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	idx_meta_description_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	idx_alt_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	idx_link_url_rw_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	idx_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	idx_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
							  " ALTER TABLE tb_menu ADD " + _
							  " 	m_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
							  " ALTER TABLE tb_menuItem ADD " + _
							  " 	mi_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	mi_link_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	mi_image_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	mi_tag_title_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
							  " ALTER TABLE tb_pagineSito ADD " + _
							  " 	id_pagDyn_" + lingua_abbr + " int NULL," + _
							  " 	id_pagStage_" + lingua_abbr + " int NULL," + _
							  " 	nome_ps_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	PAGE_keywords_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	PAGE_description_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
							  " ALTER TABLE tb_storico_index ADD " + _
							  " 	si_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	si_link_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + _
							  " ALTER TABLE tb_webs ADD " + _
							  " 	lingua_" + lingua_abbr + " bit NULL," + _
							  " 	titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
							  " 	META_keywords_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + _
							  " 	META_description_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
							  " UPDATE tb_webs SET lingua_" + lingua_abbr + " = 0;" + _
							  "	ALTER TABLE tb_webs ALTER COLUMN " + _
							  " 	lingua_" + lingua_abbr + " bit NOT NULL;"
end function


function Update_language_NextFAQ(lingua_abbr)
	Update_language_NextFAQ = " ALTER TABLE tb_FAQ ADD " + _
		  " 	faq_domanda_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	faq_risposta_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE tb_FAQ_categorie ADD " + _
		  " 	cat_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	cat_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;"
end function


function Update_language_NextLink(lingua_abbr)
	Update_language_NextLink = " ALTER TABLE tb_links ADD " + _
		  " 	link_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	link_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE tb_links_categorie ADD " + _
		  " 	cat_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;"
end function


function Update_language_NextNews(lingua_abbr)
	Update_language_NextNews = " ALTER TABLE tb_news ADD " + _
		  " 	news_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	news_estratto_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE tb_news_categorie ADD " + _
		  " 	cat_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" 
end function


function Update_language_NextTeam(lingua_abbr)
	Update_language_NextTeam = " ALTER TABLE Otb_componenti ADD " + _
		  " 	com_posizione_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	com_curriculum_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE Otb_livelli ADD " + _
		  " 	lvl_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;"
end function


function Update_language_NextGallery(lingua_abbr)
	Update_language_NextGallery = " ALTER TABLE prel_descrittori_gallery ADD " + _
		  " 	rdi_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	rdi_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE ptb_categorieGallery ADD " + _
		  " 	catC_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + _
		  " 	catC_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + _
		  " ALTER TABLE ptb_descrittori ADD " + _
		  " 	des_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + _
		  " ALTER TABLE ptb_gallery ADD " + _
		  " 	gallery_name_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + _
		  " ALTER TABLE ptb_Immagini ADD " + _
		  " 	I_Didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;"   
end function



function Aggiornamento__FRAMEWORK_CORE__146(conn, lingua_abbr)
	Aggiornamento__FRAMEWORK_CORE__146 = _
		Update_language_NextCom_NextPassport(lingua_abbr) + _
		Update_language_NextWeb(lingua_abbr) + _
		Update_language_NextFAQ(lingua_abbr) + _
		Update_language_NextLink(lingua_abbr) + _
		Update_language_NextNews(lingua_abbr) + _
		Update_language_NextTeam(lingua_abbr) + _
		Update_language_NextGallery(lingua_abbr)
end function
'*******************************************************************************************


sub AggiornamentoSpeciale__FRAMEWORK_CORE__146(conn, lingua_codice, lingua_nome_IT, lingua_nome)
	Dim sql, rs
	set rs = Server.CreateObject("ADODB.Recordset")

	sql = " SELECT * FROM tb_cnt_lingue"
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	rs.addNew
	rs("lingua_codice") = lingua_codice
	rs("lingua_nome_IT") = lingua_nome_IT
	rs("lingua_nome") = lingua_nome
	rs.update

	rs.close
	set rs = nothing
end sub


'*******************************************************************************************

function Update_language_Cancella_lingua(lingua_abbr)
	Update_language_Cancella_lingua = " ALTER TABLE rel_siti_descrittori DROP COLUMN " + _
		  " 	rsd_valore_" + lingua_abbr + " ," + _
		  " 	rsd_memo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE rel_siti_eventi DROP COLUMN " + _
		  " 	rse_email_oggetto_" + lingua_abbr + " ," + _
		  " 	rse_email_testo_" + lingua_abbr + " ," + _
		  " 	rse_sms_testo_" + lingua_abbr + " ," + _
		  " 	rse_fax_oggetto_" + lingua_abbr + " ," + _
		  " 	rse_fax_testo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_siti_descrittori DROP COLUMN " + _
		  " 	sid_nome_" + lingua_abbr + " ," + _
		  " 	sid_unita_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_siti_descrittori_raggruppamenti DROP COLUMN " + _
		  " 	sdr_titolo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_siti_eventi DROP COLUMN " + _
		  " 	sev_nome_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_siti_tabelle DROP COLUMN " + _
		  " 	tab_field_titolo_" + lingua_abbr + " ," + _
		  " 	tab_field_titolo_alt_" + lingua_abbr + " ," + _
		  " 	tab_field_descrizione_" + lingua_abbr + " ," + _
		  " 	tab_field_codice_" + lingua_abbr + " ," + _
		  " 	tab_field_url_" + lingua_abbr + " ," + _
		  " 	tab_field_meta_keywords_" + lingua_abbr + " ," + _
		  " 	tab_field_meta_description_" + lingua_abbr + " ," + _
		  " 	tab_field_return_url_" + lingua_abbr + " ," + _
		  " 	tab_tags_fields_csv_" + lingua_abbr + " ," + _
		  " 	tab_tags_fields_ssv_" + lingua_abbr + " ;" + _

		  " ALTER TABLE tb_contents DROP COLUMN " + _
		  " 	co_titolo_" + lingua_abbr + " ," + _
		  " 	co_chiave_" + lingua_abbr + " ," + _
		  " 	co_descrizione_" + lingua_abbr + " ," + _
		  " 	co_link_url_" + lingua_abbr + " ," + _
		  " 	co_meta_keywords_" + lingua_abbr + " ," + _
		  " 	co_meta_description_" + lingua_abbr + " ," + _
		  " 	co_alt_" + lingua_abbr + " ," + _
		  " 	co_link_url_rw_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_contents_index DROP COLUMN " + _
		  " 	idx_link_url_" + lingua_abbr + " ," + _
		  " 	idx_meta_keywords_" + lingua_abbr + " ," + _
		  " 	idx_meta_description_" + lingua_abbr + " ," + _
		  " 	idx_alt_" + lingua_abbr + " ," + _
		  " 	idx_link_url_rw_" + lingua_abbr + " ," + _
		  " 	idx_titolo_" + lingua_abbr + " ," + _
		  " 	idx_descrizione_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_menu DROP COLUMN " + _
		  " 	m_nome_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_menuItem DROP COLUMN " + _
		  " 	mi_titolo_" + lingua_abbr + " ," + _
		  " 	mi_link_" + lingua_abbr + " ," + _
		  " 	mi_image_" + lingua_abbr + " ," + _
		  " 	mi_tag_title_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_pagineSito DROP COLUMN " + _
		  " 	id_pagDyn_" + lingua_abbr + " ," + _
		  " 	id_pagStage_" + lingua_abbr + " ," + _
		  " 	nome_ps_" + lingua_abbr + " ," + _
		  " 	PAGE_keywords_" + lingua_abbr + " ," + _
		  " 	PAGE_description_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_storico_index DROP COLUMN " + _
		  " 	si_titolo_" + lingua_abbr + " ," + _
		  " 	si_link_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_webs DROP COLUMN " + _
		  " 	lingua_" + lingua_abbr + " ," + _
		  " 	titolo_" + lingua_abbr + " ," + _
		  " 	META_keywords_" + lingua_abbr + " ," + _
		  " 	META_description_" + lingua_abbr + " ;" + _

		  " ALTER TABLE tb_FAQ DROP COLUMN " + _
		  " 	faq_domanda_" + lingua_abbr + " ," + _
		  " 	faq_risposta_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_FAQ_categorie DROP COLUMN " + _
		  " 	cat_nome_" + lingua_abbr + " ," + _
		  " 	cat_descr_" + lingua_abbr + " ;" + _

		  " ALTER TABLE tb_links DROP COLUMN " + _
		  " 	link_nome_" + lingua_abbr + " ," + _
		  " 	link_descr_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_links_categorie DROP COLUMN " + _
		  " 	cat_nome_" + lingua_abbr + " ;" + _

		  " ALTER TABLE tb_news DROP COLUMN " + _
		  " 	news_titolo_" + lingua_abbr + " ," + _
		  " 	news_estratto_" + lingua_abbr + " ;" + _
		  " ALTER TABLE tb_news_categorie DROP COLUMN " + _
		  " 	cat_nome_" + lingua_abbr + " ;" + _

		  " ALTER TABLE Otb_componenti DROP COLUMN " + _
		  " 	com_posizione_" + lingua_abbr + " ," + _
		  " 	com_curriculum_" + lingua_abbr + " ;" + _
		  " ALTER TABLE Otb_livelli DROP COLUMN " + _
		  " 	lvl_nome_" + lingua_abbr + " ;" + _

		  " ALTER TABLE prel_descrittori_gallery DROP COLUMN " + _
		  " 	rdi_valore_" + lingua_abbr + " ," + _
		  " 	rdi_memo_" + lingua_abbr + " ;" + _
		  " ALTER TABLE ptb_categorieGallery DROP COLUMN " + _
		  " 	catC_nome_" + lingua_abbr + " ," + _
		  " 	catC_descr_" + lingua_abbr + " ;" + _
		  " ALTER TABLE ptb_descrittori DROP COLUMN " + _
		  " 	des_nome_" + lingua_abbr + " ;" + _
		  " ALTER TABLE ptb_gallery DROP COLUMN " + _
		  " 	gallery_name_" + lingua_abbr + " ;" + _
		  " ALTER TABLE ptb_Immagini DROP COLUMN " + _
		  " 	I_Didascalia_" + lingua_abbr + " ;" + _
		  		  
		  " DELETE FROM tb_cnt_lingue WHERE lingua_codice = '" + lingua_abbr + "';"
end function



function AggiornamentoSpeciale__FRAMEWORK_CORE__147(conn, lingua_abbr)
	AggiornamentoSpeciale__FRAMEWORK_CORE__147 = _
		Update_language_Cancella_lingua(lingua_abbr)
end function
'*******************************************************************************************

'*******************************************************************************************
'*******************************************************************************************
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 148
'...........................................................................................
'	Giacomo, 26/03/2010
'...........................................................................................
'   aggiunge campo per nome in inglese degli applicativi del framework
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__148(conn)
		Aggiornamento__FRAMEWORK_CORE__148 = _
				" ALTER TABLE tb_siti ADD " + _
				" 	sito_nome_en " + SQL_CharField(Conn, 250) + " NULL ;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 149
'...........................................................................................
'	Nicola, 16/04/2010
'...........................................................................................
'   aggiunge campi per foto di default su definizione contenuto dell'indice.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__149(conn)
		Aggiornamento__FRAMEWORK_CORE__149 = _
				" ALTER TABLE tb_siti_tabelle ADD " + _
				" 	tab_default_foto_thumb " + SQL_CharField(Conn, 250) + " NULL," + _
				" 	tab_default_foto_zoom " + SQL_CharField(Conn, 250) + " NULL ;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 150
'...........................................................................................
'	Nicola, 03/05/2010
'...........................................................................................
'   aggiunge campi per attivare e disattivare statistiche su indice e pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__150(conn)
		Aggiornamento__FRAMEWORK_CORE__150 = _
				" ALTER TABLE tb_webs ADD " + _
				" 	statistiche_attive bit NULL; " + _
				" UPDATE tb_webs SET statistiche_attive=1; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 151
'...........................................................................................
'	Andrea, 03/05/2010
'...........................................................................................
'   crea le viste nelle varie lingue per v_indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__151(conn)

	DropObject conn,"v_indice_it","VIEW"
	DropObject conn,"v_indice_en","VIEW"
	DropObject conn,"v_indice_es","VIEW"
	DropObject conn,"v_indice_de","VIEW"
	DropObject conn,"v_indice_cn","VIEW"
	DropObject conn,"v_indice_ru","VIEW"
	DropObject conn,"v_indice_pt","VIEW"
	DropObject conn,"v_indice_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, " + vbCrLF + _
					  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  " (" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  " ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, " + vbCrLF + _
    				  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  "(" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  "("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto, " + vbCrLF + _ 
					  "tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, " + vbCrLF + _
                      "tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	
	
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
	Agg_es = Replace(Agg_cn,"_cn","_es")
	
	Aggiornamento__FRAMEWORK_CORE__151 = _
		DropObject(conn,"v_indice_it","VIEW") + _
		DropObject(conn,"v_indice_en","VIEW") + _
		DropObject(conn,"v_indice_fr","VIEW") + _
		DropObject(conn,"v_indice_de","VIEW") + _
		DropObject(conn,"v_indice_es","VIEW") + _
		DropObject(conn,"v_indice_ru","VIEW") + _
		DropObject(conn,"v_indice_pt","VIEW") + _
		DropObject(conn,"v_indice_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__151 = Aggiornamento__FRAMEWORK_CORE__151 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 152
'...........................................................................................
'	Andrea, 03/05/2010
'...........................................................................................
'   crea le viste nelle varie lingue per v_indice_visibile
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__152(conn)

	DropObject conn,"v_indice_visibile_it","VIEW"
	DropObject conn,"v_indice_visibile_en","VIEW"
	DropObject conn,"v_indice_visibile_es","VIEW"
	DropObject conn,"v_indice_visibile_de","VIEW"
	DropObject conn,"v_indice_visibile_cn","VIEW"
	DropObject conn,"v_indice_visibile_ru","VIEW"
	DropObject conn,"v_indice_visibile_pt","VIEW"
	DropObject conn,"v_indice_visibile_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _ 
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     tb_contents.co_visibile=1 AND " + vbCrLF + _
		"          tb_contents_index.idx_visibile_assoluto=1 AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, tb_contents.co_id, " + vbCrLF + _
                      "tb_contents.co_F_table_id, tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, " + vbCrLF + _
                      "tb_contents.co_foto_thumb, tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, " + vbCrLF + _
                      "tb_contents.co_descrizione_IT, tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, " + vbCrLF + _
                      "tb_contents.co_data_scadenza, tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, " + vbCrLF + _
                      "tb_contents.co_modAdmin_id, tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, " + vbCrLF + _
                      "tb_contents.co_link_pagina_id, tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, " + vbCrLF + _
                      "tb_contents.co_meta_description_it, tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_it, tb_contents.co_link_url_rw_en, tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, " + vbCrLF + _
                      "tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, " + vbCrLF + _
                      "tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     tb_contents.co_visibile=1 AND " + vbCrLF + _
		"          tb_contents_index.idx_visibile_assoluto=1 AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__FRAMEWORK_CORE__152 = _
		DropObject(conn,"v_indice_visibile_it","VIEW") + _
		DropObject(conn,"v_indice_visibile_en","VIEW") + _
		DropObject(conn,"v_indice_visibile_fr","VIEW") + _
		DropObject(conn,"v_indice_visibile_de","VIEW") + _
		DropObject(conn,"v_indice_visibile_es","VIEW") + _
		DropObject(conn,"v_indice_visibile_ru","VIEW") + _
		DropObject(conn,"v_indice_visibile_pt","VIEW") + _
		DropObject(conn,"v_indice_visibile_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__152 = Aggiornamento__FRAMEWORK_CORE__152 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 153
'...........................................................................................
'	Giacomo, 31/05/2010
'...........................................................................................
'   aggiunge campo per la favicon
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__153(conn)
		Aggiornamento__FRAMEWORK_CORE__153 = _
				" ALTER TABLE tb_webs ADD " + _
				" 	favicon " + SQL_CharField(Conn, 255) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 154
'...........................................................................................
'	Giacomo, 01/06/2010
'...........................................................................................
'   aggiunge campo per nome tabella dalla quale recuperare i campi url
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__154(conn)
		Aggiornamento__FRAMEWORK_CORE__154 = _
				" ALTER TABLE tb_siti_tabelle ADD " + _
				" 	tab_return_url_name " + SQL_CharField(Conn, 255) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 155
'...........................................................................................
'	Giacomo, 01/06/2010
'...........................................................................................
'   crea le viste nelle varie lingue per v_indice (aggiunta campo tab_return_url_name su tb_siti_tabelle)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__155(conn)

	DropObject conn,"v_indice_it","VIEW"
	DropObject conn,"v_indice_en","VIEW"
	DropObject conn,"v_indice_es","VIEW"
	DropObject conn,"v_indice_de","VIEW"
	DropObject conn,"v_indice_cn","VIEW"
	DropObject conn,"v_indice_ru","VIEW"
	DropObject conn,"v_indice_pt","VIEW"
	DropObject conn,"v_indice_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
					  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  " (" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  " ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom,  tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
    				  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  "(" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  "("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto, " + vbCrLF + _ 
					  "tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, " + vbCrLF + _
                      "tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	
	
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
	Agg_es = Replace(Agg_cn,"_cn","_es")
	
	Aggiornamento__FRAMEWORK_CORE__155 = _
		DropObject(conn,"v_indice_it","VIEW") + _
		DropObject(conn,"v_indice_en","VIEW") + _
		DropObject(conn,"v_indice_fr","VIEW") + _
		DropObject(conn,"v_indice_de","VIEW") + _
		DropObject(conn,"v_indice_es","VIEW") + _
		DropObject(conn,"v_indice_ru","VIEW") + _
		DropObject(conn,"v_indice_pt","VIEW") + _
		DropObject(conn,"v_indice_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__155 = Aggiornamento__FRAMEWORK_CORE__155 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 156
'...........................................................................................
'	Giacomo, 01/06/2010
'...........................................................................................
'   crea le viste nelle varie lingue per v_indice_visibile (aggiunta campo tab_return_url_name su tb_siti_tabelle)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__156(conn)

	DropObject conn,"v_indice_visibile_it","VIEW"
	DropObject conn,"v_indice_visibile_en","VIEW"
	DropObject conn,"v_indice_visibile_es","VIEW"
	DropObject conn,"v_indice_visibile_de","VIEW"
	DropObject conn,"v_indice_visibile_cn","VIEW"
	DropObject conn,"v_indice_visibile_ru","VIEW"
	DropObject conn,"v_indice_visibile_pt","VIEW"
	DropObject conn,"v_indice_visibile_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _ 
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, tb_contents.co_id, " + vbCrLF + _
                      "tb_contents.co_F_table_id, tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, " + vbCrLF + _
                      "tb_contents.co_foto_thumb, tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, " + vbCrLF + _
                      "tb_contents.co_descrizione_IT, tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, " + vbCrLF + _
                      "tb_contents.co_data_scadenza, tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, " + vbCrLF + _
                      "tb_contents.co_modAdmin_id, tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, " + vbCrLF + _
                      "tb_contents.co_link_pagina_id, tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, " + vbCrLF + _
                      "tb_contents.co_meta_description_it, tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_it, tb_contents.co_link_url_rw_en, tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, " + vbCrLF + _
                      "tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, " + vbCrLF + _
                      "tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name  " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__FRAMEWORK_CORE__156 = _
		DropObject(conn,"v_indice_visibile_it","VIEW") + _
		DropObject(conn,"v_indice_visibile_en","VIEW") + _
		DropObject(conn,"v_indice_visibile_fr","VIEW") + _
		DropObject(conn,"v_indice_visibile_de","VIEW") + _
		DropObject(conn,"v_indice_visibile_es","VIEW") + _
		DropObject(conn,"v_indice_visibile_ru","VIEW") + _
		DropObject(conn,"v_indice_visibile_pt","VIEW") + _
		DropObject(conn,"v_indice_visibile_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__156 = Aggiornamento__FRAMEWORK_CORE__156 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 157
'...........................................................................................
'	Giacomo, 03/06/2010
'...........................................................................................
'   corregge il formato di alcuni campi e crea gli indici per la tabella tb_contents_index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__157(conn)
		Aggiornamento__FRAMEWORK_CORE__157 = _
				" ALTER TABLE tb_news " + _
				" ALTER COLUMN news_titolo_it " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_news " + _
				" ALTER COLUMN news_titolo_en " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_news " + _
				" ALTER COLUMN news_titolo_es " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_news " + _
				" ALTER COLUMN news_titolo_de " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_news " + _
				" ALTER COLUMN news_titolo_fr " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_FAQ " + _
				" ALTER COLUMN faq_domanda_it " + SQL_CharField(Conn, 250) + " NULL ;" + _
				" ALTER TABLE tb_FAQ " + _
				" ALTER COLUMN faq_domanda_en " + SQL_CharField(Conn, 250) + " NULL ;" + _
				" ALTER TABLE tb_FAQ " + _
				" ALTER COLUMN faq_domanda_de " + SQL_CharField(Conn, 250) + " NULL ;" + _
				" ALTER TABLE tb_FAQ " + _
				" ALTER COLUMN faq_domanda_fr " + SQL_CharField(Conn, 250) + " NULL ;" + _
				" ALTER TABLE tb_FAQ " + _
				" ALTER COLUMN faq_domanda_es " + SQL_CharField(Conn, 250) + " NULL ;" + _
				" ALTER TABLE tb_links " + _
				" ALTER COLUMN link_nome_it " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_links " + _
				" ALTER COLUMN link_nome_en " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_links " + _
				" ALTER COLUMN link_nome_de " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_links " + _
				" ALTER COLUMN link_nome_fr " + SQL_CharField(Conn, 255) + " NULL ;" + _
				" ALTER TABLE tb_links " + _
				" ALTER COLUMN link_nome_es " + SQL_CharField(Conn, 255) + " NULL ;"
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__157 = Aggiornamento__FRAMEWORK_CORE__157 + _
				" CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_CN] ON [dbo].[tb_contents_index] " + _
				"	( idx_link_url_cn ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_DE] ON [dbo].[tb_contents_index] " + _
				"	( idx_link_url_de ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_EN] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_en ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_ES] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_es ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_FR] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_fr ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_IT] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_it ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_PT] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_pt ASC ); " + _
				"CREATE NONCLUSTERED INDEX [IX_tb_contents_index_url_ru] ON [dbo].[tb_contents_index]  " + _
				"	( idx_link_url_ru ASC ); "
		end if
end function
'*******************************************************************************************
			


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 158
'...........................................................................................
'	Giacomo, 07/06/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap.
' 	imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__158(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__158 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__158 = Aggiornamento__FRAMEWORK_CORE__158 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__158 = Aggiornamento__FRAMEWORK_CORE__158 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND tab_name <> 'tb_webs' "
	else
		Aggiornamento__FRAMEWORK_CORE__158 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in LINGUE_CODICI
				Aggiornamento__FRAMEWORK_CORE__158 = Aggiornamento__FRAMEWORK_CORE__158 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"		   AND tab_name <> 'tb_webs'"
			next
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 159
'...........................................................................................
'	Giacomo, 07/06/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap.
' 	imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__159(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
			Aggiornamento__FRAMEWORK_CORE__159 = Aggiornamento__FRAMEWORK_CORE__159 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
		
					Aggiornamento__FRAMEWORK_CORE__159 = Aggiornamento__FRAMEWORK_CORE__159 + _
						"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
						"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
						"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
						"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND tab_name <> 'tb_webs' " + _
						" ; "
		next
	else
		for each lingua in LINGUE_CODICI
			Aggiornamento__FRAMEWORK_CORE__159 = Aggiornamento__FRAMEWORK_CORE__159 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    (SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"		   AND tab_name <> 'tb_webs')" + _
					" ; "
		next
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 160
'...........................................................................................
'	Giacomo, 14/06/2010
'...........................................................................................
'   aggiunge campo per il nome del database sul quale sono archiviate le e-mail
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__160(conn)
		Aggiornamento__FRAMEWORK_CORE__160 = _
				" ALTER TABLE tb_email ADD " + _
				" 	email_name_database " + SQL_CharField(Conn, 255) + " NULL ;"
				if Application("DATA_ARCHIVE_ConnectionString")<>"" then
					dim Aconn
					set Aconn = server.CreateObject("ADODB.Connection")
					Aconn.Open Application("DATA_ARCHIVE_ConnectionString"), "", "" 
					Aggiornamento__FRAMEWORK_CORE__160 = Aggiornamento__FRAMEWORK_CORE__160 + _ 
							" UPDATE tb_email SET email_name_database = '" & cString(Aconn.DefaultDatabase) & "'"
					Aconn.close
					set Aconn = nothing
				end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 161
'...........................................................................................
'	Giacomo, 06/07/2010
'...........................................................................................
'   tolgo il vincolo di integrità referenziale e lo rimetto non referenziale tra tb_webs e tb_cnt_lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__161(conn)
		Aggiornamento__FRAMEWORK_CORE__161 = _
			"ALTER TABLE  tb_webs DROP CONSTRAINT FK_tb_webs_tb_cnt_lingue"
end function
'SQL_RemoveForeignKey(conn, "tb_webs", "lingua_iniziale", "tb_cnt_lingue", true, "FK_tb_webs_tb_cnt_lingue")
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 162
'...........................................................................................
'	Nicola, 04/08/2010
'...........................................................................................
'   aggiunge parametro al next-gallery per il controllo del descrittore di tipo anagrafica
'	con relativo filtro di ricerca
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__162(conn)
	Aggiornamento__FRAMEWORK_CORE__162 = SQL_AddForeignKey(conn, "tb_webs", "lingua_iniziale", "tb_cnt_lingue", "lingua_codice", false, "")
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__162(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTGALLERY))>0 then
		CALL AddParametroSito(conn, "CARATTERISTICHE_CONTATTI_ABILITATE", _
									null, _
									"Abilita il descrittore di tipo anagrafiche", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTGALLERY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "CARATTERISTICHE_CONTATTI_RUBRICA", _
									null, _
									"Attiva il filtro di selezione delle anagrafiche solo per una rubrica", _
									"", _
									adIDispatch, _
									0, _
									"", _
									1, _
									1, _
									NEXTGALLERY, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 163
'...........................................................................................
'	Andrea, 23/09/2010
'...........................................................................................
'   crea la tabella per il log del request checker
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__163(conn)
	Aggiornamento__FRAMEWORK_CORE__163 = " create table " & SQL_dbo(conn) & "log_request_checker("+_
		"log_id " + SQL_PrimaryKey(conn,"log_request_checker")+" , "+_
		"log_date datetime NULL, "+_
		"log_parameter_name NVARCHAR(100) NULL, "+_
		"log_parameter_value NVARCHAR(100) NULL, "+_
		"log_rawhttp NTEXT NULL, "+_
		"log_url NTEXT NULL); "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 164
'...........................................................................................
'	Sergio, 29/09/2010
'...........................................................................................
'   aggiunge campi mancanti in alcuni db su tb_contents_index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__164(conn)
 dim rs
 
 set rs = server.CreateObject("ADODB.recordset")
 sql = "SELECT TOP 1 * FROM tb_contents_index" 
 rs.open sql, conn, adOpenDynamic, adLockOptimistic
        
 if not FieldExists(rs,"idx_titolo_it") then
	Aggiornamento__FRAMEWORK_CORE__164 = "  " + _
		Update_language_Content_Index("it") + _
		Update_language_Content_Index("en") + _
		Update_language_Content_Index("fr") + _
		Update_language_Content_Index("de") + _
		Update_language_Content_Index("es") + _
	  " "
 else
	Aggiornamento__FRAMEWORK_CORE__164 = "SELECT * FROM AA_versione"
 end if
end function


' Usata da Aggiornamento__FRAMEWORK_CORE__164
function Update_language_Content_Index(lingua_abbr)
	Update_language_Content_Index = " ALTER TABLE tb_contents_index ADD " + _
		  " 	idx_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + _
		  " 	idx_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 165
'...........................................................................................
'	Matteo, 13/10/2010
'...........................................................................................
'   crea la tabella per rss
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__165(conn)
	Aggiornamento__FRAMEWORK_CORE__165 = " CREATE TABLE " & SQL_dbo(conn) & "tb_rss(" +_
		"rss_id "+SQL_PrimaryKey(conn,"tb_rss")+" , " +_
		"rss_web_id INT NOT NULL, " +_
		"rss_lingua NVARCHAR(10) NOT NULL, " +_
		"rss_file NVARCHAR(250) NOT NULL, " +_
		"rss_titolo NVARCHAR(250) NULL, " +_
		"rss_descrizione NVARCHAR(250) NULL, " +_
		"rss_image NVARCHAR(250) NULL, " +_
		"rss_query NTEXT NULL, " +_
		"rss_freq_generazione INT NULL, " +_
		"rss_data_generazione DATETIME NULL); " +_
		SQL_AddForeignKey(conn, "tb_rss", "rss_web_id", "tb_webs", "id_webs", false, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 166
'...........................................................................................
'	Matteo, 20/10/2010
'...........................................................................................
'   aggiunge i flag per l'abilitazione e la gestione metatag su tb_rss
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__166(conn)
	Aggiornamento__FRAMEWORK_CORE__166 = _
		" ALTER TABLE " & SQL_dbo(conn) & "tb_rss ADD " + _
		" 		rss_abilitato BIT NULL, " + _
		" 		rss_metatag BIT NULL; " + _
		" UPDATE tb_rss " + _
		"	 SET rss_abilitato = 0, " + _
		" 		 rss_metatag = 0; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 167
'...........................................................................................
'	Andrea, 2/11/2010
'...........................................................................................
'   crea la tabella per le query e i relativi parametri
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__167(conn)
	Aggiornamento__FRAMEWORK_CORE__167 = " create table " & SQL_dbo(conn) & "tb_query("+_
		"query_id " + SQL_PrimaryKey(conn,"tb_query")+" , "+_
		"query_nome NVARCHAR(250) NULL, "+_
		"query_descrizione NTEXT NULL, "+_
		"query_code NTEXT NULL, "+_
		"query_id_app int NULL); "
	Aggiornamento__FRAMEWORK_CORE__167 = Aggiornamento__FRAMEWORK_CORE__167 + " create table " & SQL_dbo(conn) & "tb_query_params("+_
		"param_id " + SQL_PrimaryKey(conn,"tb_query_params")+" , "+_
		"param_query_id int NULL, "+_
		"param_nome NVARCHAR(150) NULL, "+_
		"param_tipo NVARCHAR(50) NULL, "+_
		"param_ordine int NULL, "+_
		"param_desc NTEXT NULL); " +_		
		SQL_AddForeignKey(conn, "tb_query_params", "param_query_id", "tb_query", "query_id", false, "")
end function

'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 168
'...........................................................................................
'	Andrea, 8/11/2010
'...........................................................................................
'   aggiunta campo sottoquery a tb_query_params
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__168(conn)
	Aggiornamento__FRAMEWORK_CORE__168 = " alter table "& SQL_dbo(conn) & "tb_query_params " +_
		"add param_sottoquery NTEXT;"
end function

'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 169
'...........................................................................................
'	Andrea, 17/11/2010
'...........................................................................................
'   aggiunta campi a tb_siti_tabelle_pubblicazioni
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__169(conn)
	Aggiornamento__FRAMEWORK_CORE__169 = " ALTER TABLE "& SQL_dbo(conn) & "tb_siti_tabelle_pubblicazioni " +_
		"ADD pub_field_foto_thumb " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_foto_zoom " & SQL_CharField(Conn, 255) & " null, "+_
		"pub_field_titolo_it " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_en " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_fr " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_de " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_es " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_it " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_en " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_fr " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_de " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_es" & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_it " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_en " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_fr " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_de " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_es " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_meta_keywords_it " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_en " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_fr " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_de " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_es " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_it" & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_en " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_fr " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_de " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_es " & SQL_CharField(Conn, 0) & "  null;"
	if(DB_Type(conn)=DB_SQL) then
		Aggiornamento__FRAMEWORK_CORE__169 = Aggiornamento__FRAMEWORK_CORE__169 + " ALTER TABLE "& SQL_dbo(conn) & "tb_siti_tabelle_pubblicazioni " +_
		"ADD pub_field_titolo_ru " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_cn " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_pt " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_ru " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_cn " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_titolo_alt_pt " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_ru " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_cn " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_descrizione_pt " & SQL_CharField(Conn, 255) & "  null, "+_
		"pub_field_meta_keywords_ru " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_cn " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_keywords_pt " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_ru " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_cn " & SQL_CharField(Conn, 0) & "  null, "+_
		"pub_field_meta_description_pt " & SQL_CharField(Conn, 0) & "  null;"
	end if	
end function

'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 170
'...........................................................................................
'	Giacomo, 23/11/2010
'...........................................................................................
'	aggiunge campo per decidere se indicizzare (pagina per pagina e intero sito)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__170(conn)
	Aggiornamento__FRAMEWORK_CORE__170 = _
		" ALTER TABLE tb_pagineSito ADD " + _
		"	indicizzabile BIT NULL; " + _
		" UPDATE tb_pagineSito SET indicizzabile = 1 ; " + _
		" ALTER TABLE tb_pagineSito ALTER COLUMN indicizzabile BIT NOT NULL; " + _
		" ALTER TABLE tb_webs ADD " + _
		"	sito_indicizzabile BIT NULL; " + _
		" UPDATE tb_webs SET sito_indicizzabile = 1 ; " + _
		" ALTER TABLE tb_webs ALTER COLUMN sito_indicizzabile BIT NOT NULL; " + _
		" DELETE FROM tb_webs_metatag WHERE meta_name LIKE 'robots' ; " 
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 171
'...........................................................................................
'	Giacomo, 23/11/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili.
' 	imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__171(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
			Aggiornamento__FRAMEWORK_CORE__171 = Aggiornamento__FRAMEWORK_CORE__171 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
		
					Aggiornamento__FRAMEWORK_CORE__171 = Aggiornamento__FRAMEWORK_CORE__171 + _
						"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
						"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
						"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
						"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND tab_name <> 'tb_webs' AND indicizzabile " + _
						" ; "
		next
	else
		for each lingua in LINGUE_CODICI
			Aggiornamento__FRAMEWORK_CORE__171 = Aggiornamento__FRAMEWORK_CORE__171 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    (SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND tab_name <> 'tb_webs')" + _
					" ; "
		next
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 172
'...........................................................................................
'	Giacomo, 24/11/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili.
' 	imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__172(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__172 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__172 = Aggiornamento__FRAMEWORK_CORE__172 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__172 = Aggiornamento__FRAMEWORK_CORE__172 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE (idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND tab_name <> 'tb_webs' AND indicizzabile "
	else
		Aggiornamento__FRAMEWORK_CORE__172 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in LINGUE_CODICI
				Aggiornamento__FRAMEWORK_CORE__172 = Aggiornamento__FRAMEWORK_CORE__172 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND tab_name <> 'tb_webs'"
			next
	end if
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 173
'...........................................................................................
'	Giacomo, 03/12/2010
'...........................................................................................
'   Aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili
'	oppure escludendo tutti gli elementi se l'intero sito non è indicizzabile.
'	Aggiunto inoltre filtro per far camparire sempre la root, ma mai due volte.
' 	Imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__173(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
			Aggiornamento__FRAMEWORK_CORE__173 = Aggiornamento__FRAMEWORK_CORE__173 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
		
					Aggiornamento__FRAMEWORK_CORE__173 = Aggiornamento__FRAMEWORK_CORE__173 + _
						"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
						"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
						"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
						"    	 WHERE sito_indicizzabile AND ((idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND indicizzabile " + _
						"				AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)) " + _
						" ; "
		next
	else
		for each lingua in LINGUE_CODICI
			Aggiornamento__FRAMEWORK_CORE__173 = Aggiornamento__FRAMEWORK_CORE__173 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS " + _
					"    (SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (" & SQL_IsTrue(conn, "sito_indicizzabile") & ") AND ((idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)))" + _
					" ; "
		next
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 174
'...........................................................................................
'	Giacomo, 03/12/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili
'	oppure escludendo tutti gli elementi se l'intero sito non è indicizzabile.
'	Aggiunto inoltre filtro per far camparire sempre la root, ma mai due volte.
' 	Imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__174(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__174 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__174 = Aggiornamento__FRAMEWORK_CORE__174 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__174 = Aggiornamento__FRAMEWORK_CORE__174 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE sito_indicizzabile AND ((idx_principale OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND indicizzabile " + _
				"				AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)) "
	else
		Aggiornamento__FRAMEWORK_CORE__174 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in LINGUE_CODICI
				Aggiornamento__FRAMEWORK_CORE__174 = Aggiornamento__FRAMEWORK_CORE__174 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (" & SQL_IsTrue(conn, "sito_indicizzabile") & ") AND ((idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id))"
			next
	end if
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 175
'...........................................................................................
'	Giacomo, 06/12/2010
'...........................................................................................
'   aggiunge parametro al next-web 5 per permettere di copiare pagine e templates anche 
'	da un sito ad un'altro
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__175(conn)
	Aggiornamento__FRAMEWORK_CORE__175 = " SELECT * FROM AA_Versione "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__175(conn)
	if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM tb_siti WHERE id_sito = " & NEXTWEB5)) > 0 then
		CALL AddParametroSito(conn, "COPIA_PAGINE_TRA_SITI", _
									null, _
									"Permette di copiare pagine e templates anche da un sito ad un'altro", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTWEB5, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 176
'...........................................................................................
'	Giacomo, 07/12/2010
'...........................................................................................
'   Aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili
'	oppure escludendo tutti gli elementi se l'intero sito non è indicizzabile.
'	Aggiunto inoltre filtro per far camparire sempre la root, ma mai due volte.
' 	Imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'	...ulteriore correzione (OR idx_livello = 0)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__176(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
			Aggiornamento__FRAMEWORK_CORE__176 = Aggiornamento__FRAMEWORK_CORE__176 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS "  + vbCrLf + _
					"    SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
		
					Aggiornamento__FRAMEWORK_CORE__176 = Aggiornamento__FRAMEWORK_CORE__176 + _
						"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
						"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
						"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
						"    	 WHERE sito_indicizzabile AND ((idx_principale OR idx_livello = 0 OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND indicizzabile " + _
						"				AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)) " + _
						" ; "
		next
	else
		for each lingua in LINGUE_CODICI
			Aggiornamento__FRAMEWORK_CORE__176 = Aggiornamento__FRAMEWORK_CORE__176 + _
					DropObject(conn, "v_indice_sitemap_" + lingua, "VIEW") + _
					" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap_" + lingua + " AS "  + vbCrLf + _
					"    (SELECT DISTINCT " + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (" & SQL_IsTrue(conn, "sito_indicizzabile") & ") AND ((idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR idx_livello = 0 OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)))" + _
					" ; "
		next
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 177
'...........................................................................................
'	Giacomo, 07/12/2010
'...........................................................................................
'   aggiorna vista sitemap per esclusione elementi non marcati come inclusi nelle sitemap e come non indicizzabili
'	oppure escludendo tutti gli elementi se l'intero sito non è indicizzabile.
'	Aggiunto inoltre filtro per far camparire sempre la root, ma mai due volte.
' 	Imposta flag di visibilità su sitemap per tutti i contenuti.
'   Per ACCESS solo 5 lingue, per SQL tutte e 8 le lingue
'	...ulteriore correzione (OR idx_livello = 0)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__177(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__177 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__177 = Aggiornamento__FRAMEWORK_CORE__177 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__177 = Aggiornamento__FRAMEWORK_CORE__177 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE sito_indicizzabile AND ((idx_principale OR idx_livello = 0 OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND indicizzabile " + _
				"				AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)) "
	else
		Aggiornamento__FRAMEWORK_CORE__177 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in LINGUE_CODICI
				Aggiornamento__FRAMEWORK_CORE__177 = Aggiornamento__FRAMEWORK_CORE__177 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
					"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
					"    WHERE (" & SQL_IsTrue(conn, "sito_indicizzabile") & ") AND ((idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR idx_livello = 0 OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & ")" + vbcRLF + _
					"		   AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id))"
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 178
'...........................................................................................
'	Giacomo, 17/01/2011
'...........................................................................................
'   cambia il tipo della colonna PraticaPrefisso da nvarchar(5) a nvarchar(50)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__178(conn)
	Aggiornamento__FRAMEWORK_CORE__178 = _
		" ALTER TABLE tb_Indirizzario ALTER COLUMN PraticaPrefisso " & SQL_CharField(Conn, 50) & " null "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 179
'...........................................................................................
'	Andrea, 3/02/2011
'...........................................................................................
'   crea indici per ottimizzazione indice e pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__179(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__179 = _
				" CREATE NONCLUSTERED INDEX [IX_tb_contents_index_padre_id] ON [dbo].[tb_contents_index] " + vbcRLF + _
				" ( [idx_padre_id] ASC ) ON [PRIMARY]; " + vbcRLF + _
				" CREATE NONCLUSTERED INDEX [IX_tb_contents_index_tipologie_padre_lista] ON [dbo].[tb_contents_index]  " + vbcRLF + _
				" ( idx_tipologie_padre_lista ASC ) ON [PRIMARY]; " + vbcRLF + _
				" CREATE NONCLUSTERED INDEX [IX_tb_layers_id_pag] ON [dbo].tb_layers " + vbcRLF + _
				" ( id_pag ASC ) ON [PRIMARY]; " + vbcRLF + _
				" CREATE NONCLUSTERED INDEX [IX_tb_contents_index_principale] ON [dbo].[tb_contents_index] " + vbcRLF + _
				" ( idx_principale DESC ) ON [PRIMARY]; "
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__179 = _
				" CREATE INDEX [IX_tb_contents_index_padre_id] ON [tb_contents_index] " + vbcRLF + _
				" ( [idx_padre_id] ASC ); " + vbcRLF + _
				" CREATE INDEX [IX_tb_contents_index_tipologie_padre_lista] ON [tb_contents_index]  " + vbcRLF + _
				" ( idx_tipologie_padre_lista ASC ); " + vbcRLF + _
				" CREATE INDEX [IX_tb_layers_id_pag] ON tb_layers " + vbcRLF + _
				" ( id_pag ASC ); " + vbcRLF + _
				" CREATE INDEX [IX_tb_contents_index_principale] ON [tb_contents_index] " + vbcRLF + _
				" ( idx_principale DESC ); "					
	end select		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 180
'...........................................................................................
'	Andrea, 3/02/2011
'...........................................................................................
'   crea trigger solo per sql server per il salvataggio ed il recupero degli url nella cancellazione delle voci alternative.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__180(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__180 = _			
				" -- ============================================= " + vbcRLF + _
				" -- Author:		Andrea " + vbcRLF + _
				" -- Create date: 03-02-2011 " + vbcRLF + _
				" -- Description:	Si attiva alla cancellazione di " + vbcRLF + _
				" --				un nodo dell'indice non principale " + vbcRLF + _
				" --				e ne salva gli url associandoli " + vbcRLF + _
				" --				al corrispondente nodo principale " + vbcRLF + _
				" -- ============================================= " + vbcRLF + _
				" CREATE TRIGGER [dbo].[tb_contents_index_delete] " + vbcRLF + _
				" ON  tb_contents_index " + vbcRLF + _
				" AFTER DELETE " + vbcRLF + _
				" AS " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				" DECLARE @idx_id_deleted int " + vbcRLF + _
				" DECLARE @is_principale bit " + vbcRLF + _
				" DECLARE @idx_content_deleted int " + vbcRLF + _
				" -- Creo un cursore per delete multipli da utilizzare come recordset sulle righe eliminate " + vbcRLF + _
				" DECLARE rs CURSOR local FAST_FORWARD FOR SELECT idx_id,idx_principale,idx_content_id FROM deleted " + vbcRLF + _
				" OPEN rs " + vbcRLF + _
				" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
				" WHILE @@FETCH_STATUS = 0 " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				" --SELECT @idx_id_deleted=idx_id,@is_principale=idx_principale,@idx_content_deleted=idx_content_id FROM deleted " + vbcRLF + _
				" IF @is_principale=0 " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				"	-- Recupero l'idx_id del nodo principale " + vbcRLF + _
				"	DECLARE @idx_principale int		" + vbcRLF + _
				"	SELECT TOP 1 @idx_principale=idx_id FROM tb_contents_index WHERE idx_content_id=@idx_content_deleted ORDER BY idx_principale DESC "	+ vbcRLF + _		 
				"	-- Associo tutti gli url dello storico al nodo principale " + vbcRLF + _
				"	UPDATE rel_index_url_redirect SET riu_idx_id=@idx_principale, riu_modData=GETDATE() WHERE riu_idx_id=@idx_id_deleted " + vbcRLF + _
				"	-- Recupero l'id admin da inserire nello storico " + vbcRLF + _
				"	DECLARE @id_admin int " + vbcRLF + _
				"	SELECT TOP 1 @id_admin=riu_insAdmin_id FROM rel_index_url_redirect " + vbcRLF + _
				"	-- Recupero gli url attivi per ogni lingua e se diversi da "" gli inserisco nello storico " + vbcRLF + _
				"	DECLARE @url_attivo VARCHAR(500) " + vbcRLF + _
				"	-- ITALIANO " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_it,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'it',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- INGLESE " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_en,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'en',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- FRANCESE " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_fr,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'fr',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- TEDESCO " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_de,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'de',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- SPAGNOLO " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_es,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'es',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END	"	 + vbcRLF + _
				"	-- RUSSO " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_ru,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'ru',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- CINESE " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_cn,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'cn',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				"	-- PORTOGHESE " + vbcRLF + _
				"	SELECT @url_attivo=IsNull(idx_link_url_rw_pt,'') FROM deleted " + vbcRLF + _
				"	IF @url_attivo<>'' " + vbcRLF + _
				"	BEGIN " + vbcRLF + _
				"		INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'pt',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"	END " + vbcRLF + _
				" END " + vbcRLF + _
				" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
				" END " + vbcRLF + _
				" DEALLOCATE rs " + vbcRLF + _
				" END " 
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__180 = "SELECT * FROM AA_Versione"
					
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 181
'...........................................................................................
'	Nicola, 04/02/2011
'...........................................................................................
'   corregge vista v_indice_visibile perchè non funziona con access 2010
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__181(conn)
	Aggiornamento__FRAMEWORK_CORE__181 = _
        DropObject(conn, "v_indice_visibile", "VIEW") + _
		"CREATE VIEW " & SQL_Dbo(Conn) & "v_indice_visibile AS " + vbCrLF + _
		"    SELECT * FROM (tb_contents_index INNER JOIN tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id ) " + vbCrLF + _
		"                  INNER JOIN tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"    WHERE " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 182
'...........................................................................................
'	Nicola, 04/02/2011
'...........................................................................................
'   corregge viste v_indice_visibile_<lingua> perchè non funziona con access 2010
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__182(conn)

	DropObject conn,"v_indice_visibile_it","VIEW"
	DropObject conn,"v_indice_visibile_en","VIEW"
	DropObject conn,"v_indice_visibile_es","VIEW"
	DropObject conn,"v_indice_visibile_de","VIEW"
	DropObject conn,"v_indice_visibile_cn","VIEW"
	DropObject conn,"v_indice_visibile_ru","VIEW"
	DropObject conn,"v_indice_visibile_pt","VIEW"
	DropObject conn,"v_indice_visibile_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _ 
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, tb_contents.co_id, " + vbCrLF + _
                      "tb_contents.co_F_table_id, tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, " + vbCrLF + _
                      "tb_contents.co_foto_thumb, tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, " + vbCrLF + _
                      "tb_contents.co_descrizione_IT, tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, " + vbCrLF + _
                      "tb_contents.co_data_scadenza, tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, " + vbCrLF + _
                      "tb_contents.co_modAdmin_id, tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, " + vbCrLF + _
                      "tb_contents.co_link_pagina_id, tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, " + vbCrLF + _
                      "tb_contents.co_meta_description_it, tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_it, tb_contents.co_link_url_rw_en, tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, " + vbCrLF + _
                      "tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, " + vbCrLF + _
                      "tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name  " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          (tb_contents.co_data_pubblicazione>= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          (tb_contents.co_data_scadenza<= " & SQL_now(conn) & " OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__FRAMEWORK_CORE__182 = _
		DropObject(conn,"v_indice_visibile_it","VIEW") + _
		DropObject(conn,"v_indice_visibile_en","VIEW") + _
		DropObject(conn,"v_indice_visibile_fr","VIEW") + _
		DropObject(conn,"v_indice_visibile_de","VIEW") + _
		DropObject(conn,"v_indice_visibile_es","VIEW") + _
		DropObject(conn,"v_indice_visibile_ru","VIEW") + _
		DropObject(conn,"v_indice_visibile_pt","VIEW") + _
		DropObject(conn,"v_indice_visibile_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__182 = Aggiornamento__FRAMEWORK_CORE__182 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 183
'...........................................................................................
'	Giacomo, 07/02/2011
'...........................................................................................
'   correzione su trigger per il salvataggio ed il recupero degli url nella cancellazione delle voci alternative.
'	(aggiunto controllo per verificare che idx_principale sia maggiore di 0)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__183(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__183 = _
				DropObject(conn, "tb_contents_index_delete", "TRIGGER") + vbcRLF + _
				" -- ============================================= " + vbcRLF + _
				" -- Author:		Andrea " + vbcRLF + _
				" -- Create date: 03-02-2011 " + vbcRLF + _
				" -- Description:	Si attiva alla cancellazione di " + vbcRLF + _
				" --				un nodo dell'indice non principale " + vbcRLF + _
				" --				e ne salva gli url associandoli " + vbcRLF + _
				" --				al corrispondente nodo principale " + vbcRLF + _
				" -- ============================================= " + vbcRLF + _
				" CREATE TRIGGER [dbo].[tb_contents_index_delete] " + vbcRLF + _
				" ON  tb_contents_index " + vbcRLF + _
				" AFTER DELETE " + vbcRLF + _
				" AS " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				" DECLARE @idx_id_deleted int " + vbcRLF + _
				" DECLARE @is_principale bit " + vbcRLF + _
				" DECLARE @idx_content_deleted int " + vbcRLF + _
				" -- Creo un cursore per delete multipli da utilizzare come recordset sulle righe eliminate " + vbcRLF + _
				" DECLARE rs CURSOR local FAST_FORWARD FOR SELECT idx_id,idx_principale,idx_content_id FROM deleted " + vbcRLF + _
				" OPEN rs " + vbcRLF + _
				" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
				" WHILE @@FETCH_STATUS = 0 " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				" --SELECT @idx_id_deleted=idx_id,@is_principale=idx_principale,@idx_content_deleted=idx_content_id FROM deleted " + vbcRLF + _
				" IF @is_principale=0 " + vbcRLF + _
				" BEGIN " + vbcRLF + _
				"	-- Recupero l'idx_id del nodo principale " + vbcRLF + _
				"	DECLARE @idx_principale int		" + vbcRLF + _
				"	SELECT TOP 1 @idx_principale=idx_id FROM tb_contents_index WHERE idx_content_id=@idx_content_deleted ORDER BY idx_principale DESC "	+ vbcRLF + _		 
				"	-- Associo tutti gli url dello storico al nodo principale " + vbcRLF + _
				"	IF ISNULL(@idx_principale,0) > 0 " + vbcRLF + _
				"		BEGIN " + vbcRLF + _
				"			UPDATE rel_index_url_redirect SET riu_idx_id=@idx_principale, riu_modData=GETDATE() WHERE riu_idx_id=@idx_id_deleted " + vbcRLF + _
				"			-- Recupero l'id admin da inserire nello storico " + vbcRLF + _
				"			DECLARE @id_admin int " + vbcRLF + _
				"			SELECT TOP 1 @id_admin=riu_insAdmin_id FROM rel_index_url_redirect " + vbcRLF + _
				"			-- Recupero gli url attivi per ogni lingua e se diversi da "" gli inserisco nello storico " + vbcRLF + _
				"			DECLARE @url_attivo VARCHAR(500) " + vbcRLF + _
				"			-- ITALIANO " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_it,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'it',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- INGLESE " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_en,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'en',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- FRANCESE " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_fr,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'fr',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- TEDESCO " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_de,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'de',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- SPAGNOLO " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_es,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'es',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END	"	 + vbcRLF + _
				"			-- RUSSO " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_ru,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'ru',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- CINESE " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_cn,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'cn',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"			-- PORTOGHESE " + vbcRLF + _
				"			SELECT @url_attivo=IsNull(idx_link_url_rw_pt,'') FROM deleted " + vbcRLF + _
				"			IF @url_attivo<>'' " + vbcRLF + _
				"			BEGIN " + vbcRLF + _
				"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'pt',GETDATE(),@id_admin,GETDATE(),@id_admin) " + vbcRLF + _
				"			END " + vbcRLF + _
				"		END " + vbcRLF + _
				" END " + vbcRLF + _
				" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
				" END " + vbcRLF + _
				" DEALLOCATE rs " + vbcRLF + _
				" END " 
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__183 = "SELECT * FROM AA_Versione"
					
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 184
'...........................................................................................
'	Andrea, 03/03/2011
'...........................................................................................
'   creazione tabella per gestione tag da query
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__184(conn)
	Aggiornamento__FRAMEWORK_CORE__184 = _
	"CREATE TABLE " & SQL_dbo(conn) & "tb_siti_tabelle_tag_query("+_
	"tq_id " + SQL_PrimaryKey(conn,"tb_siti_tabelle_tag_query")+" , "+_
	"tq_tab_id int, "+_
	"tq_nome NVARCHAR(100) NULL, "+_
	"tq_query NTEXT NULL, "+_
	"tq_separatore NVARCHAR(10) NULL);" + vbCrLF + _
	SQL_AddForeignKey(conn, "tb_siti_tabelle_tag_query", "tq_tab_id", "tb_siti_tabelle", "tab_id", true, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 185
'...........................................................................................
'	Andrea, 03/03/2011
'...........................................................................................
'   creazione tabella per gestione log del sistema
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__185(conn)
	Aggiornamento__FRAMEWORK_CORE__185 = _
	"CREATE TABLE " & SQL_dbo(conn) & "log_framework("+_
	"log_id " + SQL_PrimaryKey(conn,"log_framework")+" , "+_
	"log_table_nome NVARCHAR(50) NULL, "+_
	"log_record_id int, "+_
	"log_codice NVARCHAR(50) NULL, "+_
	"log_descrizione NVARCHAR(255) NULL, "+_
	"log_data smalldatetime, "+_
	"log_admin_id int, "+_
	"log_user_id int, "+_
	"log_http_request NTEXT NULL, "+_
	"log_application_id int);" + vbCrLF + _
	SQL_AddForeignKey(conn, "log_framework", "log_admin_id", "tb_admin", "id_admin", false, "")+ vbCrLF + _
	SQL_AddForeignKey(conn, "log_framework", "log_user_id", "tb_Utenti", "ut_id", false, "")+ vbCrLF + _
	SQL_AddForeignKey(conn, "log_framework", "log_application_id", "tb_siti", "id_sito", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 186
'...........................................................................................
'	Giacomo, 15/03/2011
'...........................................................................................
'   aggiunta colonna su tb_siti per modifica gestione permessi per applicazione
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__186(conn)
	Aggiornamento__FRAMEWORK_CORE__186 = _
		" ALTER TABLE tb_siti ADD " & _
		"	sito_protetto bit NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 187
'...........................................................................................
'	Matteo, 16/03/2011
'...........................................................................................
'   aggiunge funzione SQL per coalesce personalizzato sulle lingue (esclude anche la stringa vuota)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__187(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__187 = _
			" CREATE FUNCTION dbo.fn_next_coalesce " + vbCrLf + _
			" ( " + vbCrLf + _
			" 	@valore1 nvarchar(500), " + vbCrLf + _
			" 	@valore2 nvarchar(500), " + vbCrLf + _
			" 	@valore3 nvarchar(500) " + vbCrLf + _
			" ) " + vbCrLf + _
			" RETURNS nvarchar(500) " + vbCrLf + _
			" AS " + vbCrLf + _
			" BEGIN " + vbCrLf + _
			" 	DECLARE @result nvarchar(500) " + vbCrLf + _
			" 	IF (COALESCE(@valore1, '') <> '') " + vbCrLf + _
			" 	  SET @result = @valore1  " + vbCrLf + _
			" 	ELSE IF (COALESCE(@valore2, '') <> '') " + vbCrLf + _
			" 	  SET @result = @valore2 " + vbCrLf + _
			" 	ELSE IF (COALESCE(@valore3, '') <> '') " + vbCrLf + _
			" 	  SET @result = @valore3 " + vbCrLf + _
			" 	ELSE " + vbCrLf + _
			" 	  SET @result = '' " + vbCrLf + _	
			" 	RETURN @result " + vbCrLf + _
			" END " 
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__187 = "SELECT * FROM AA_Versione"
	end select
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 188
'...........................................................................................
'	Andrea , 18/03/2011
'...........................................................................................
'   aggiunte colonne per gestione plugin tipo html
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__188(conn)
	Aggiornamento__FRAMEWORK_CORE__188 = _
		" ALTER TABLE tb_objects ADD " + vbCrLf + _
		"	obj_html_it NTEXT NULL, "  + vbCrLf + _
		"	obj_html_en NTEXT NULL, "  + vbCrLf + _
		"	obj_html_fr NTEXT NULL, "  + vbCrLf + _
		"	obj_html_de NTEXT NULL, "  + vbCrLf + _
		"	obj_html_es NTEXT NULL " 
	Select case DB_Type(conn)		
		case DB_SQL		
			Aggiornamento__FRAMEWORK_CORE__188 = Aggiornamento__FRAMEWORK_CORE__188  + vbCrLf + _
				",	obj_html_ru NTEXT NULL, "  + vbCrLf + _
				"	obj_html_cn NTEXT NULL, "  + vbCrLf + _
				"	obj_html_pt NTEXT NULL; "
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__188 = Aggiornamento__FRAMEWORK_CORE__188 & ";"	
	end select			
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 189
'...........................................................................................
'	Andrea , 26/04/2011
'...........................................................................................
'   aggiunge campo per distinguere siti mobili
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__189(conn)
	Aggiornamento__FRAMEWORK_CORE__189 = _
		" ALTER TABLE tb_webs ADD " + vbCrLf + _
		" sito_mobile bit default 0 not null, "  + vbCrLf + _
		" URL_alternativo NTEXT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 190
'...........................................................................................
'	Andrea, 04/05/2011
'...........................................................................................
'   creazione tabella sentinella per tb_contents_index
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__190(conn)
	Aggiornamento__FRAMEWORK_CORE__190 = _
	"CREATE TABLE " & SQL_dbo(conn) & "tb_contents_index_sentinel("+_
	"sent_time datetime);" +_
	" INSERT INTO tb_contents_index_sentinel (sent_time) VALUES "
	Select case DB_Type(conn)		
		case DB_SQL		
			Aggiornamento__FRAMEWORK_CORE__190 = Aggiornamento__FRAMEWORK_CORE__190  +  "(GETDATE());"
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__190 = Aggiornamento__FRAMEWORK_CORE__190  +  "(NOW());"
	end select
	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 191
'...........................................................................................
'	Giacomo, 09/05/2011
'...........................................................................................
'   aggiunta colonna su tb_utenti con riferimento a tb_admin
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__191(conn)
	Aggiornamento__FRAMEWORK_CORE__191 = _
		" ALTER TABLE tb_Utenti ADD " & _
		"	ut_admin_id int NULL; " & _
		SQL_AddForeignKey(conn, "tb_Utenti", "ut_admin_id", "tb_admin", "id_admin", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 192
'...........................................................................................
'	Giacomo, 11/05/2011
'...........................................................................................
'   aggiunta colonne su tb_Indirizzario
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__192(conn)
	Aggiornamento__FRAMEWORK_CORE__192 = _
		" ALTER TABLE tb_Indirizzario ADD " & _
			" 	cnt_insAdmin_id int NULL; " & _
		" ALTER TABLE tb_Indirizzario ADD " & _
			" 	cnt_insData smalldatetime NULL; " & _
		" ALTER TABLE tb_Indirizzario ADD " & _
			" 	cnt_modAdmin_id int NULL; " & _
		" ALTER TABLE tb_Indirizzario ADD " & _
			" 	cnt_modData smalldatetime NULL; " & _
		SQL_AddForeignKey(conn, "tb_Indirizzario", "cnt_insAdmin_id", "tb_admin", "ID_admin", false, "") & _
		SQL_AddForeignKey(conn, "tb_Indirizzario", "cnt_modAdmin_id", "tb_admin", "ID_admin", false, "_2")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 193
'...........................................................................................
'	Giacomo, 12/05/2011
'...........................................................................................
'   aggiunta colonna su tb_admin per salvare la richiesta http raw
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__193(conn)
	Aggiornamento__FRAMEWORK_CORE__193 = _
		" ALTER TABLE log_admin ADD " & _
			" 	log_http_raw " + SQL_CharField(Conn, 0) + " NULL ;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 194
'...........................................................................................
'	Giacomo, 12/05/2011
'...........................................................................................
'  	aggiornamento per criptare tutte le password di tb_admin
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__194(conn)
	Aggiornamento__FRAMEWORK_CORE__194 = "SELECT * FROM AA_Versione "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__194(conn)
	dim password
	sql = "SELECT * FROM tb_admin"
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		password = EncryptPassword(rs("admin_password"))
		rs("admin_password") = password
		rs.update
		rs.moveNext
	wend
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 195
'...........................................................................................
'	Giacomo, 19/05/2011
'...........................................................................................
'   crea indici per ottimizzazione indice e pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__195(conn)
	Select case DB_Type(conn)		
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__195 = _
				" CREATE INDEX [IX_tb_utenti] ON [dbo].[tb_Utenti] " + vbcRLF + _
				" (" + vbcRLF + _
				"	[ut_NextCom_ID] ASC " + vbcRLF + _
				" ); "
		case DB_Access
			Aggiornamento__FRAMEWORK_CORE__195 = _
				" CREATE INDEX [IX_tb_utenti] ON [tb_Utenti] " + vbcRLF + _
				" (" + vbcRLF + _
				"	[ut_NextCom_ID] ASC " + vbcRLF + _
				" ); "				
	end select		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 196
'...........................................................................................
'	Giacomo, 12/05/2011
'...........................................................................................
'   
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__196(conn)
	Aggiornamento__FRAMEWORK_CORE__196 = _
		" ALTER TABLE tb_siti_tabelle ADD " & _
			" 	tab_field_return_foto_thumb " + SQL_CharField(Conn, 255) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 197
'...........................................................................................
' Giacomo 01/06/2011
'...........................................................................................
' Creazione tabelle per la nuova gestione delle newsletter 
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__197(conn)
    Aggiornamento__FRAMEWORK_CORE__197 = _
			"CREATE TABLE " + SQL_Dbo(conn) + "tb_newsletters(" + _
			"	nl_id " + SQL_PrimaryKey(conn, "tb_newsletters") + ", " + _
			"	nl_nome_it " + SQL_CharField(Conn, 255) + " NULL, " + _
			"	nl_pagina_id int NULL," + _
			" 	nl_insAdmin_id int NULL, " + _
			" 	nl_insData DATETIME NULL, " + _
			" 	nl_modAdmin_id int NULL, " + _
			" 	nl_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "tb_newsletters", "nl_pagina_id", "tb_pagineSito", "id_pagineSito", false, "") + _
			SQL_AddForeignKey(conn, "tb_newsletters", "nl_insAdmin_id", "tb_admin", "ID_admin", false, "_1") + _
			SQL_AddForeignKey(conn, "tb_newsletters", "nl_modAdmin_id", "tb_admin", "ID_admin", false, "_2") + _
			"CREATE TABLE " + SQL_Dbo(conn) + "tb_newsletters_contents(" + _
			"	nlc_id " + SQL_PrimaryKey(conn, "tb_newsletters_contents") + ", " + _
			"	nlc_idx_id int NULL, " + _
			"	nlc_tipo_id int NULL, " + _
			"	nlc_data_invio smalldatetime NULL, " + _
			"	nlc_email_inviata_id int NULL, " + _
			" 	nlc_insAdmin_id int NULL, " + _
			" 	nlc_insData DATETIME NULL, " + _
			" 	nlc_modAdmin_id int NULL, " + _
			" 	nlc_modData DATETIME NULL " + _
			"); " + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_idx_id", "tb_contents_index", "idx_id", false, "") + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_tipo_id", "tb_newsletters", "nl_id", false, "") + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_email_inviata_id", "tb_email", "email_id", false, "") + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_insAdmin_id", "tb_admin", "ID_admin", false, "_1") + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_modAdmin_id", "tb_admin", "ID_admin", false, "_2")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 198
'...........................................................................................
' Giacomo 13/06/2011 
'...........................................................................................
' Correzioni su tabelle newsletter
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__198(conn)
    Aggiornamento__FRAMEWORK_CORE__198 = _
			SQL_RemoveForeignKey(conn, "tb_newsletters_contents", "nlc_idx_id", "tb_contents_index", false, "") + _
			"ALTER TABLE tb_newsletters_contents DROP COLUMN nlc_idx_id; " + _
			"ALTER TABLE tb_newsletters_contents ADD nlc_co_id int NULL; " + _
			SQL_AddForeignKey(conn, "tb_newsletters_contents", "nlc_co_id", "tb_contents", "co_id", false, "") + _
			"ALTER TABLE tb_newsletters ADD nl_lingua " + SQL_CharField(Conn, 2) + " NULL ;" + _
			"ALTER TABLE tb_newsletters_contents ADD nlc_ordine int NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 199
'...........................................................................................
' Giacomo 24/06/2011 
'...........................................................................................
' aggiungo colonna su tb_email
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__199(conn)
    Aggiornamento__FRAMEWORK_CORE__199 = _
			"ALTER TABLE tb_email ADD email_newsletter_tipo_id int NULL; " + _
			SQL_AddForeignKey(conn, "tb_email", "email_newsletter_tipo_id", "tb_newsletters", "nl_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 200
'...........................................................................................
' Nicola 28/06/2011 
'...........................................................................................
' aggiungo colonna per embed video su news
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__200(conn)
    Aggiornamento__FRAMEWORK_CORE__200 = _
			"ALTER TABLE tb_news " + _
			" ADD news_html_embed " + SQL_CharField(Conn, 0) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 201
'...........................................................................................
'	Giacomo, 29/06/2011
'...........................................................................................
'   corregge viste v_indice_visibile_<lingua> per correggere le condizioni della data di pubblicazione e scadenza
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__201(conn)

	DropObject conn,"v_indice_visibile_it","VIEW"
	DropObject conn,"v_indice_visibile_en","VIEW"
	DropObject conn,"v_indice_visibile_es","VIEW"
	DropObject conn,"v_indice_visibile_de","VIEW"
	DropObject conn,"v_indice_visibile_cn","VIEW"
	DropObject conn,"v_indice_visibile_ru","VIEW"
	DropObject conn,"v_indice_visibile_pt","VIEW"
	DropObject conn,"v_indice_visibile_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _ 
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, tb_contents.co_id, " + vbCrLF + _
                      "tb_contents.co_F_table_id, tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, " + vbCrLF + _
                      "tb_contents.co_foto_thumb, tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, " + vbCrLF + _
                      "tb_contents.co_descrizione_IT, tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, " + vbCrLF + _
                      "tb_contents.co_data_scadenza, tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, " + vbCrLF + _
                      "tb_contents.co_modAdmin_id, tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, " + vbCrLF + _
                      "tb_contents.co_link_pagina_id, tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, " + vbCrLF + _
                      "tb_contents.co_meta_description_it, tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_it, tb_contents.co_link_url_rw_en, tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, " + vbCrLF + _
                      "tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, " + vbCrLF + _
                      "tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name  " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__FRAMEWORK_CORE__201 = _
		DropObject(conn,"v_indice_visibile_it","VIEW") + _
		DropObject(conn,"v_indice_visibile_en","VIEW") + _
		DropObject(conn,"v_indice_visibile_fr","VIEW") + _
		DropObject(conn,"v_indice_visibile_de","VIEW") + _
		DropObject(conn,"v_indice_visibile_es","VIEW") + _
		DropObject(conn,"v_indice_visibile_ru","VIEW") + _
		DropObject(conn,"v_indice_visibile_pt","VIEW") + _
		DropObject(conn,"v_indice_visibile_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__201 = Aggiornamento__FRAMEWORK_CORE__201 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 202
'...........................................................................................
' Giacomo 11/07/2011
' aggiunta colonne lingue mancanti su tb_tipNumeri
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__202(conn, lingua_abbr)
	Aggiornamento__FRAMEWORK_CORE__202 = _
		  " ALTER TABLE tb_tipNumeri ADD " + vbCrLf + _
		  " 	nome_tiponumero_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 203
'...........................................................................................
'	Giacomo, 14/07/2011
'...........................................................................................
'   aggiunge parametro al next-passport
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__203(conn)
	Aggiornamento__FRAMEWORK_CORE__203 = "SELECT * FROM AA_versione"
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__203(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTPASSPORT))>0 then
		CALL AddParametroSito(conn, "PAGINA_AVVISO_ABILITAZIONE_UTENTE", _
									0, _
									"Pagina da inviare dopo aver abilitato un utente.", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTPASSPORT, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 204
'...........................................................................................
'	Giacomo, 04/08/2011
'...........................................................................................
'  	aggiornamento per aggiungere il codiceInserimento ai contatti sprovvisti
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__204(conn)
	Aggiornamento__FRAMEWORK_CORE__204 = "SELECT * FROM AA_Versione "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__204(conn)
	dim password, rs
	if DB_Type(conn) = DB_SQL then
		sql = " SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE ISNULL(codiceInserimento,'')='' "
	else
		sql = " SELECT IDElencoIndirizzi FROM tb_indirizzario WHERE ISNULL(codiceInserimento) OR codiceInserimento = '' "
	end if
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
	while not rs.eof
		CALL SetCodiceInserimento(conn, rs("IDElencoIndirizzi"))
		rs.moveNext
	wend
	rs.close
	set rs = nothing
end sub
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 205
'...........................................................................................
'	Nicola, 20/09/2011
'...........................................................................................
'   rimuove indici non ottimizzati e crea indici per ottimizzazione indice e pagine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__205(conn)
	Aggiornamento__FRAMEWORK_CORE__205 = DropObject(conn, "IX_tb_contents_index_principale", "INDEX")
	Select case DB_Type(conn)
		case DB_SQL					
			Aggiornamento__FRAMEWORK_CORE__205 = Aggiornamento__FRAMEWORK_CORE__205 + _
				DropObject(conn, "IX_tb_contents_index_url_IT", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_EN", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_FR", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_DE", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_ES", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_PT", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_CN", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_url_RU", "INDEX") + _
				DropObject(conn, "IDX_tb_contents_tabella_contenuto", "INDEX") + _
				DropObject(conn, "IDX_tb_contents_index_urls", "INDEX") + _
				DropObject(conn, "IX_tb_contents_co_F_ids", "INDEX") + _
				DropObject(conn, "IX_tb_contents_index_urls", "INDEX") + _
				" CREATE NONCLUSTERED INDEX IX_tb_contents_co_F_ids ON tb_contents " + _
							 " ( co_F_table_id ASC, co_F_key_id ASC ) ; " + _
				" CREATE NONCLUSTERED INDEX IX_tb_contents_index_urls ON tb_contents_index " + _
							 " ( idx_link_url_IT ASC, " + _
							   " idx_link_url_EN ASC, " + _
							   " idx_link_url_FR ASC, " + _
							   " idx_link_url_DE ASC, " + _
							   " idx_link_url_ES ASC, " + _
							   " idx_link_url_RU ASC, " + _
							   " idx_link_url_CN ASC, " + _
							   " idx_link_url_PT ASC ); "
		case DB_Access
			
	end select		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 206
'...........................................................................................
' Giacomo 03/10/2011
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__206(conn)
	Aggiornamento__FRAMEWORK_CORE__206 = _
			" ALTER TABLE tb_admin " + _
			" ALTER COLUMN admin_cognome " + SQL_CharField(Conn, 255) + " NULL ;" + _
			" ALTER TABLE tb_admin " + _
			" ALTER COLUMN admin_nome " + SQL_CharField(Conn, 255) + " NULL ;" + _
			" ALTER TABLE tb_indirizzario " + _
			" ALTER COLUMN CognomeElencoIndirizzi " + SQL_CharField(Conn, 255) + " NULL ;" + _
			" ALTER TABLE tb_indirizzario " + _
			" ALTER COLUMN SecondoNomeElencoIndirizzi " + SQL_CharField(Conn, 255) + " NULL ;" + _
			" ALTER TABLE tb_indirizzario " + _
			" ALTER COLUMN NomeElencoIndirizzi " + SQL_CharField(Conn, 255) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 207
'...........................................................................................
' Giacomo 29/12/2011
'...........................................................................................
' Aggiungo campo nome in lingua alle rubriche
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__207(conn)
	Aggiornamento__FRAMEWORK_CORE__207 = _
			" ALTER TABLE tb_rubriche ADD " + _
			SQL_MultiLanguageFieldComplete(conn, " nome_pubblico_rubrica_<lingua> " + SQL_CharField(Conn, 500)) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 208
'...........................................................................................
' Nicola 04/04/2012
'...........................................................................................
' Aggiungo campo su sito per aggiunta codice in cima alla pagina
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__208(conn)
	Aggiornamento__FRAMEWORK_CORE__208 = _
			" ALTER TABLE tb_webs ADD pagehead_script " + SQL_CharField(Conn, 0) + " NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 209
'...........................................................................................
'	Giacomo, 13/04/2012
'...........................................................................................
'   aggiunge parametro al com e colonna a tb_ValoriNumeri (gestione email per newsletter)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__209(conn)
	Aggiornamento__FRAMEWORK_CORE__209 = _
			" ALTER TABLE tb_ValoriNumeri " + SQL_AddColumn(conn) + _
			" 	email_newsletter bit NULL "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__209(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTCOM))>0 then
		CALL AddParametroSito(conn, "ATTIVA_RECAPITI_NEWSLETTER", _
									0, _
									"Attiva la possibilità di selezionare, per ogni contatto, l'e-mail per le spedizioni delle newsletter.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTCOM, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 210
'...........................................................................................
'	Giacomo, 07/05/2012
'...........................................................................................
'   corregge le viste nelle varie lingue per v_indice (aggiunta campi idx_titolo_ e idx_descrizione_)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__210(conn)

	DropObject conn,"v_indice_it","VIEW"
	DropObject conn,"v_indice_en","VIEW"
	DropObject conn,"v_indice_es","VIEW"
	DropObject conn,"v_indice_de","VIEW"
	DropObject conn,"v_indice_cn","VIEW"
	DropObject conn,"v_indice_ru","VIEW"
	DropObject conn,"v_indice_pt","VIEW"
	DropObject conn,"v_indice_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
					  "tb_contents_index.idx_titolo_IT, tb_contents_index.idx_titolo_EN, tb_contents_index.idx_descrizione_IT, tb_contents_index.idx_descrizione_EN, " + vbCrLF + _
					  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  " (" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  " ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom,  tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
					  "tb_contents_index.idx_titolo_IT, tb_contents_index.idx_titolo_EN, tb_contents_index.idx_descrizione_IT, tb_contents_index.idx_descrizione_EN, " + vbCrLF + _
    				  " (" + SQL_IF(conn, SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
				 	  SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLF + _
					  "(" & SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLF + _
					  "("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & ") ", 1, 0) + ") AS visibile_assoluto, " + vbCrLF + _ 
					  "tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, " + vbCrLF + _
                      "tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_contents_index.idx_titolo_cn, tb_contents_index.idx_descrizione_cn " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id;"
	
	
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
	Agg_es = Replace(Agg_cn,"_cn","_es")
	
	Aggiornamento__FRAMEWORK_CORE__210 = _
		DropObject(conn,"v_indice_it","VIEW") + _
		DropObject(conn,"v_indice_en","VIEW") + _
		DropObject(conn,"v_indice_fr","VIEW") + _
		DropObject(conn,"v_indice_de","VIEW") + _
		DropObject(conn,"v_indice_es","VIEW") + _
		DropObject(conn,"v_indice_ru","VIEW") + _
		DropObject(conn,"v_indice_pt","VIEW") + _
		DropObject(conn,"v_indice_cn","VIEW") + _
		Agg_it + Agg_en + Agg_es + Agg_fr + Agg_de 
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__210 = Aggiornamento__FRAMEWORK_CORE__210 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 211
'...........................................................................................
'	Giacomo, 07/05/2012
'...........................................................................................
'   corregge viste v_indice_visibile_<lingua> (aggiunta campi idx_titolo_ e idx_descrizione_)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__211(conn)

	DropObject conn,"v_indice_visibile_it","VIEW"
	DropObject conn,"v_indice_visibile_en","VIEW"
	DropObject conn,"v_indice_visibile_es","VIEW"
	DropObject conn,"v_indice_visibile_de","VIEW"
	DropObject conn,"v_indice_visibile_cn","VIEW"
	DropObject conn,"v_indice_visibile_ru","VIEW"
	DropObject conn,"v_indice_visibile_pt","VIEW"
	DropObject conn,"v_indice_visibile_fr","VIEW"
	
	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents.co_id, tb_contents.co_F_table_id, " + vbCrLF + _
                      "tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, tb_contents.co_foto_thumb, " + vbCrLF + _
                      "tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, tb_contents.co_descrizione_IT, " + vbCrLF + _
                      "tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, tb_contents.co_data_scadenza, " + vbCrLF + _
                      "tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, tb_contents.co_modAdmin_id, " + vbCrLF + _
                      "tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, tb_contents.co_link_pagina_id, " + vbCrLF + _
                      "tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, tb_contents.co_meta_description_it, " + vbCrLF + _
                      "tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, tb_contents.co_link_url_rw_it, " + vbCrLF + _ 
                      "tb_contents.co_link_url_rw_en, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
					  "tb_contents_index.idx_titolo_IT, tb_contents_index.idx_titolo_EN, tb_contents_index.idx_descrizione_IT, tb_contents_index.idx_descrizione_EN " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "v_indice_visibile_cn AS " + vbCrLf + _
		"SELECT     tb_contents_index.idx_id, tb_contents_index.idx_content_id, tb_contents_index.idx_link_url_IT, tb_contents_index.idx_link_url_EN, " + vbCrLF + _
                      "tb_contents_index.idx_link_tipo, tb_contents_index.idx_foglia, tb_contents_index.idx_livello, tb_contents_index.idx_ordine_assoluto, " + vbCrLF + _
                      "tb_contents_index.idx_visibile_assoluto, tb_contents_index.idx_padre_id, tb_contents_index.idx_tipologia_padre_base, " + vbCrLF + _
                      "tb_contents_index.idx_tipologie_padre_lista, tb_contents_index.idx_insData, tb_contents_index.idx_insAdmin_id, " + vbCrLF + _
                      "tb_contents_index.idx_modData, tb_contents_index.idx_modAdmin_id, tb_contents_index.idx_link_pagina_id, " + vbCrLF + _
                      "tb_contents_index.idx_ordine, tb_contents_index.idx_foto_thumb, tb_contents_index.idx_foto_zoom, " + vbCrLF + _
                      "tb_contents_index.idx_autopubblicato, tb_contents_index.idx_contatore, tb_contents_index.idx_contUtenti, " + vbCrLF + _
                      "tb_contents_index.idx_contCrawler, tb_contents_index.idx_contAltro, tb_contents_index.idx_contRes, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_it, tb_contents_index.idx_meta_keywords_en, tb_contents_index.idx_meta_description_it, " + vbCrLF + _
                      "tb_contents_index.idx_meta_description_en, tb_contents_index.idx_alt_it, tb_contents_index.idx_alt_en, " + vbCrLF + _
                      "tb_contents_index.idx_principale, tb_contents_index.idx_link_url_rw_it, tb_contents_index.idx_link_url_rw_en, " + vbCrLF + _
                      "tb_contents_index.idx_webs_id, tb_contents_index.idx_priorita, tb_contents_index.idx_link_url_cn, " + vbCrLF + _
                      "tb_contents_index.idx_meta_keywords_cn, tb_contents_index.idx_meta_description_cn, tb_contents_index.idx_alt_cn, " + vbCrLF + _
                      "tb_contents_index.idx_link_url_rw_cn, tb_contents.co_id, " + vbCrLF + _
                      "tb_contents.co_F_table_id, tb_contents.co_F_key_id, tb_contents.co_titolo_IT, tb_contents.co_titolo_EN, tb_contents.co_ordine, " + vbCrLF + _
                      "tb_contents.co_foto_thumb, tb_contents.co_foto_zoom, tb_contents.co_chiave_IT, tb_contents.co_chiave_EN, " + vbCrLF + _
                      "tb_contents.co_descrizione_IT, tb_contents.co_descrizione_EN, tb_contents.co_visibile, tb_contents.co_data_pubblicazione, " + vbCrLF + _
                      "tb_contents.co_data_scadenza, tb_contents.co_insData, tb_contents.co_insAdmin_id, tb_contents.co_modData, " + vbCrLF + _
                      "tb_contents.co_modAdmin_id, tb_contents.co_link_tipo, tb_contents.co_link_url_IT, tb_contents.co_link_url_EN, " + vbCrLF + _
                      "tb_contents.co_link_pagina_id, tb_contents.co_meta_keywords_it, tb_contents.co_meta_keywords_en, " + vbCrLF + _
                      "tb_contents.co_meta_description_it, tb_contents.co_meta_description_en, tb_contents.co_alt_it, tb_contents.co_alt_en, " + vbCrLF + _
                      "tb_contents.co_link_url_rw_it, tb_contents.co_link_url_rw_en, tb_contents.co_titolo_cn, tb_contents.co_chiave_cn, " + vbCrLF + _
                      "tb_contents.co_descrizione_cn, tb_contents.co_link_url_cn, tb_contents.co_meta_keywords_cn, tb_contents.co_meta_description_cn, " + vbCrLF + _
                      "tb_contents.co_alt_cn, tb_contents.co_link_url_rw_cn, tb_siti_tabelle.tab_id, tb_siti_tabelle.tab_sito_id, tb_siti_tabelle.tab_titolo, " + vbCrLF + _
                      "tb_siti_tabelle.tab_name, tb_siti_tabelle.tab_field_chiave, tb_siti_tabelle.tab_field_foto_thumb, tb_siti_tabelle.tab_field_foto_zoom, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_visibile, tb_siti_tabelle.tab_field_ordine, tb_siti_tabelle.tab_field_data_pubblicazione, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_data_scadenza, tb_siti_tabelle.tab_from_sql, tb_siti_tabelle.tab_colore, tb_siti_tabelle.tab_parametro, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_it, tb_siti_tabelle.tab_field_titolo_en, tb_siti_tabelle.tab_field_titolo_alt_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_alt_en, tb_siti_tabelle.tab_field_descrizione_it, tb_siti_tabelle.tab_field_descrizione_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_it, tb_siti_tabelle.tab_field_codice_en, tb_siti_tabelle.tab_field_url_it, tb_siti_tabelle.tab_field_url_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_keywords_it, tb_siti_tabelle.tab_field_meta_keywords_en, tb_siti_tabelle.tab_field_meta_description_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_en, tb_siti_tabelle.tab_thumb, tb_siti_tabelle.tab_zoom, tb_siti_tabelle.tab_tags_abilitati, " + vbCrLF + _
                      "tb_siti_tabelle.tab_ricercabile, tb_siti_tabelle.tab_per_sitemap, tb_siti_tabelle.tab_priorita_base, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_return_url_it, tb_siti_tabelle.tab_field_return_url_en, tb_siti_tabelle.tab_tags_fields_csv_it, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_csv_en, tb_siti_tabelle.tab_tags_fields_ssv_it, tb_siti_tabelle.tab_tags_fields_ssv_en, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_titolo_cn, tb_siti_tabelle.tab_field_titolo_alt_cn, tb_siti_tabelle.tab_field_descrizione_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_codice_cn, tb_siti_tabelle.tab_field_url_cn, tb_siti_tabelle.tab_field_meta_keywords_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_field_meta_description_cn, tb_siti_tabelle.tab_field_return_url_cn, tb_siti_tabelle.tab_tags_fields_csv_cn, " + vbCrLF + _
                      "tb_siti_tabelle.tab_tags_fields_ssv_cn, tb_siti_tabelle.tab_default_foto_thumb, tb_siti_tabelle.tab_default_foto_zoom, tb_siti_tabelle.tab_return_url_name, " + vbCrLF + _
					  "tb_contents_index.idx_titolo_IT, tb_contents_index.idx_titolo_EN, tb_contents_index.idx_descrizione_IT, tb_contents_index.idx_descrizione_EN, " + vbCrLF + _
					  "tb_contents_index.idx_titolo_cn, tb_contents_index.idx_descrizione_cn " + vbCrLF + _
		"FROM         (tb_contents_index INNER JOIN " + vbCrLF + _
                      "tb_contents ON tb_contents_index.idx_content_id = tb_contents.co_id) INNER JOIN " + vbCrLF + _
                      "tb_siti_tabelle ON tb_contents.co_F_table_id = tb_siti_tabelle.tab_id " + vbCrLF + _
		"WHERE     " & SQL_IsTrue(conn, "tb_contents.co_visibile") & " AND " + vbCrLF + _
		"          " & SQL_IsTrue(conn, "tb_contents_index.idx_visibile_assoluto") & " AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_pubblicazione", SQL_now(conn)) &" >= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_pubblicazione") & ") AND " + vbCrLf + _
		"          ("& SQL_DateDiff(conn, "d", "tb_contents.co_data_scadenza", SQL_now(conn)) &" <= 0 OR " & SQL_IsNull(conn, "tb_contents.co_data_scadenza") & "); "
	
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")
			
	Aggiornamento__FRAMEWORK_CORE__211 = _
		DropObject(conn,"v_indice_visibile_it","VIEW") + _
		DropObject(conn,"v_indice_visibile_en","VIEW") + _
		DropObject(conn,"v_indice_visibile_fr","VIEW") + _
		DropObject(conn,"v_indice_visibile_de","VIEW") + _
		DropObject(conn,"v_indice_visibile_es","VIEW") + _
		DropObject(conn,"v_indice_visibile_ru","VIEW") + _
		DropObject(conn,"v_indice_visibile_pt","VIEW") + _
		DropObject(conn,"v_indice_visibile_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es
		if DB_Type(conn) = DB_SQL then
			Aggiornamento__FRAMEWORK_CORE__211 = Aggiornamento__FRAMEWORK_CORE__211 + Agg_ru + Agg_pt + Agg_cn
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 212
'...........................................................................................
' Giacomo 30/05/2012 
'...........................................................................................
' aggiungo colonne su tb_newsletters
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__212(conn)
    Aggiornamento__FRAMEWORK_CORE__212 = _
			"ALTER TABLE tb_newsletters ADD " + _
			" nl_rubriche_default " + SQL_CharField(Conn, 0) + " NULL, " + _
			" nl_contatti_default " + SQL_CharField(Conn, 0) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 213
'...........................................................................................
' Giacomo 14/06/2012 
'...........................................................................................
' aggiungo colonne su tabella descrittori gallery
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__213(conn)
    Aggiornamento__FRAMEWORK_CORE__213 = _
			"ALTER TABLE ptb_descrittori ADD " + _
			" des_ordine int NULL, " + _
			" des_pub_server_id int NULL, " + _
			" des_per_ricerca bit NULL, " + _
			" des_raggruppamento_id int NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "des_unita_<lingua> " + SQL_CharField(Conn, 100)) + "," + _
			" des_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
			" des_per_confronto bit NULL, " + _
			" des_img " + SQL_CharField(Conn, 255) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 214
'...........................................................................................
' Nicola 22/06/2012 
'...........................................................................................
' aggiungo colonna su tabella amministratori per gestione email per newsletter diversa
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__214(conn)
    Aggiornamento__FRAMEWORK_CORE__214 = _
			"ALTER TABLE tb_admin ADD " + _
			" admin_email_newsletter " + SQL_CharField(Conn, 255) + " NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 215
'...........................................................................................
' Giacomo 09/08/2012 
'...........................................................................................
'aggiunge campo logo alle categorie di gallery
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__215(conn)
    Aggiornamento__FRAMEWORK_CORE__215 = _
        " ALTER TABLE ptb_categorieGallery ADD catC_logo " + SQL_CharField(Conn, 250) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 216
'...........................................................................................
' Giacomo 27/08/2012 
'...........................................................................................
'aggiunge campo su tabelle per gestire indicizzazione
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__216(conn)
    Aggiornamento__FRAMEWORK_CORE__216 = _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	tab_indicizza_per_visibilita bit NULL; " + _
		" UPDATE tb_siti_tabelle SET tab_indicizza_per_visibilita = 0; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 217
'...........................................................................................
'	Nicola, 10/09/2012
'...........................................................................................
'   aggiunge parametro al Gallery per gestione pagine con catalogo sfogliabile.
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__217(conn)
	Aggiornamento__FRAMEWORK_CORE__217 = _
			" SELECT * FROM AA_VERSIONE "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__217(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTGALLERY))>0 then
		CALL AddParametroSito(conn, "ATTIVA_GALLERY_SFOGLIABILE_CATEGORIE_IDS", _
									0, _
									"Attiva la generazione del catalogo sfogbliabile per le gallery delle categorie indicate.", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTGALLERY, _
									null, null, null, null, null)
		CALL AddParametroSito(conn, "ATTIVA_GALLERY_SFOGLIABILE_DIRECTORY", _
									0, _
									"Directory di base per la generazione del catalogo sfogliabile dalla gallery di immagini.", _
									"", _
									adVarBinary, _
									0, _
									"", _
									1, _
									1, _
									NEXTGALLERY, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 218
'...........................................................................................
' Giacomo 16/10/2012 
'...........................................................................................
'aggiunge campo foto aggiuntivo sul Next-team - scheda componente
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__218(conn)
    Aggiornamento__FRAMEWORK_CORE__218 = _
		" ALTER TABLE otb_componenti ADD " + _
		"	com_foto_zoom " + SQL_CharField(Conn, 500) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 219
'...........................................................................................
' Giacomo 18/10/2012 
'...........................................................................................
'aggiunge campo su tabelle per gestire pubblicazioni "semi-automatiche"
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__219(conn)
    Aggiornamento__FRAMEWORK_CORE__219 = _
		" ALTER TABLE tb_siti_tabelle ADD " + _
		"	tab_pagina_default_id int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 220
'...........................................................................................
' Giacomo 19/11/2012
'...........................................................................................
' Creazione categorie e descrittori per le anagrafiche del Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__220(conn)
    Aggiornamento__FRAMEWORK_CORE__220 = _
			" ALTER TABLE tb_indirizzario ADD cnt_categoria_id int NULL; " + _
			" " + _
			" CREATE TABLE " + SQL_Dbo(conn) + "itb_indirizzario_categorie (" + _
			"	icat_id  " & SQL_PrimaryKey(conn, "itb_indirizzario_categorie") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "icat_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			"	icat_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
			"	icat_foglia BIT NULL ," + _
			"	icat_livello INTEGER NULL ," + _
			"	icat_padre_id INTEGER NULL ," + _
			"	icat_ordine INTEGER NULL ," + _
			"	icat_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL ," + _				
				SQL_MultiLanguageFieldComplete(conn, "icat_descr_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			"	icat_tipologia_padre_base INTEGER NULL ," + _
			"	icat_visibile BIT NULL , " + _
			"	icat_albero_visibile BIT NULL , " + _
			"	icat_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL" + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "tb_indirizzario", "cnt_categoria_id", "itb_indirizzario_categorie", "icat_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + " tb_indirizzario_carattech (" + _
			" ict_id " & SQL_PrimaryKey(conn, "tb_indirizzario_carattech") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "ict_nome_<lingua> " + SQL_CharField(Conn, 510) + " NULL ") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "ict_unita_<lingua> " + SQL_CharField(Conn, 100) + " NULL ") + ", " + _
			" ict_tipo int NULL, " + _
			" ict_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
			" ict_per_ricerca bit NULL, " + _
			" ict_per_confronto bit NULL, " + _
			" ict_img " + SQL_CharField(Conn, 255) + " NULL, " + _
			" ict_raggruppamento_id int NULL " + _
			"); " + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "tb_indirizzario_carattech_raggruppamenti (" + _
			" icr_id " & SQL_PrimaryKey(conn, "tb_indirizzario_carattech_raggruppamenti") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "icr_titolo_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			" icr_ordine int NULL, " + _
			" icr_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
			" icr_di_sistema bit NULL " + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "tb_indirizzario_carattech", "ict_raggruppamento_id", "tb_indirizzario_carattech_raggruppamenti", "icr_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + " irel_categ_ctech (" + _
			" rcc_id " & SQL_PrimaryKey(conn, "irel_categ_ctech") + ", " + _
			" rcc_ctech_id int NULL, " + _
			" rcc_ordine int NULL, " + _
			" rcc_categoria_id int NULL " + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "irel_categ_ctech", "rcc_ctech_id", "tb_indirizzario_carattech", "ict_id", false, "") + _
			SQL_AddForeignKey(conn, "irel_categ_ctech", "rcc_categoria_id", "itb_indirizzario_categorie", "icat_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + " irel_cnt_ctech (" + _
			" ric_id " & SQL_PrimaryKey(conn, "irel_cnt_ctech") + ", " + _
			" ric_cnt_id int NULL, " + _
			" ric_ctech_id int NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "ric_valore_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "irel_cnt_ctech", "ric_cnt_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + _
			SQL_AddForeignKey(conn, "irel_cnt_ctech", "ric_ctech_id", "tb_indirizzario_carattech", "ict_id", false, "")
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 221
'...........................................................................................
' Nicola 19/11/2012
'...........................................................................................
'aggiunge tabella per gestione parco macchine - fotocopiatrici
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__221(conn)
    Aggiornamento__FRAMEWORK_CORE__221 = _
		" CREATE TABLE " & SQL_Dbo(Conn) & "tb_indirizzario_macchine (" + vbCrLf + _
		"	ima_id " & SQL_PrimaryKey(conn, "tb_indirizzario_macchine") + ", " + vbCrLf + _
		"	ima_contatto_id INT NOT NULL, " + vbCrLF + _
		"	ima_marchio " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_modello " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_numero " + SQL_CharField(Conn, 50) + " NULL," + vbCrLf + _
		"	ima_tipocolore " + SQL_CharField(Conn, 50) + " NULL," + vbCrLf + _
		"	ima_contratto " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_installazione " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_scadenza " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_matricola " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		"	ima_fornitore " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		AddInsModFields("ima") + vbCrLf + _
		" ); " + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_indirizzario_macchine", "ima_contatto_id", "tb_indirizzario", "IdElencoindirizzi", true, "")
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__221(conn)
	CALL AddParametroSito(conn, "ATTIVA_PARCO_MACCHINE", _
						0, _
						"Attiva la sezione di gestione del parco macchine del cliente.", _
						"", _
						adBoolean, _
						0, _
						"", _
						1, _
						1, _
						NEXTCOM, _
						null, null, null, null, null)
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 222
'...........................................................................................
' Giacomo 19/11/2012
'...........................................................................................
' Aggiungo parametro per attivazione categorie nel Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__222(conn)
    Aggiornamento__FRAMEWORK_CORE__222 = _
				" SELECT * FROM AA_VERSIONE "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__222(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTCOM))>0 then
		CALL AddParametroSito(conn, "NEXTCOM_ATTIVA_GESTIONE_CATEGORIE", _
							0, _
							"Attiva la sezione per la gestione delle categorie delle anagrafiche.", _
							"", _
							adBoolean, _
							0, _
							"", _
							1, _
							1, _
							NEXTCOM, _
							null, null, null, null, null)
	end if
end sub
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 223
'...........................................................................................
' Giacomo 21/11/2012
'...........................................................................................
' RI-Creazione (con nomi di tabella corretti) categorie e descrittori per le anagrafiche del Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__223(conn)
    Aggiornamento__FRAMEWORK_CORE__223 = _
			" CREATE TABLE " + SQL_Dbo(conn) + "tb_indirizzario_categorie (" + _
			"	icat_id  " & SQL_PrimaryKey(conn, "tb_indirizzario_categorie") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "icat_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			"	icat_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
			"	icat_foglia BIT NULL ," + _
			"	icat_livello INTEGER NULL ," + _
			"	icat_padre_id INTEGER NULL ," + _
			"	icat_ordine INTEGER NULL ," + _
			"	icat_ordine_assoluto " + SQL_CharField(Conn, 255) + " NULL ," + _				
				SQL_MultiLanguageFieldComplete(conn, "icat_descr_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			"	icat_tipologia_padre_base INTEGER NULL ," + _
			"	icat_visibile BIT NULL , " + _
			"	icat_albero_visibile BIT NULL , " + _
			"	icat_tipologie_padre_lista " + SQL_CharField(Conn, 255) + " NULL" + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "tb_indirizzario", "cnt_categoria_id", "tb_indirizzario_categorie", "icat_id", false, "") + ";" + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "rel_categ_ctech (" + _
			" rcc_id " & SQL_PrimaryKey(conn, "rel_categ_ctech") + ", " + _
			" rcc_ctech_id int NULL, " + _
			" rcc_ordine int NULL, " + _
			" rcc_categoria_id int NULL " + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "rel_categ_ctech", "rcc_ctech_id", "tb_indirizzario_carattech", "ict_id", false, "") + ";" + _
			SQL_AddForeignKey(conn, "rel_categ_ctech", "rcc_categoria_id", "tb_indirizzario_categorie", "icat_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "rel_cnt_ctech (" + _
			" ric_id " & SQL_PrimaryKey(conn, "rel_cnt_ctech") + ", " + _
			" ric_cnt_id int NULL, " + _
			" ric_ctech_id int NULL, " + _
			SQL_MultiLanguageFieldComplete(conn, "ric_valore_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
			"); " + _
			" " + _
			SQL_AddForeignKey(conn, "rel_cnt_ctech", "ric_cnt_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + ";" + _
			SQL_AddForeignKey(conn, "rel_cnt_ctech", "ric_ctech_id", "tb_indirizzario_carattech", "ict_id", false, "")
			
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT tb_indirizzario_categorie ON ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			" INSERT INTO tb_indirizzario_categorie (icat_id, "&SQL_MultiLanguageFieldComplete(conn, "icat_nome_<lingua>")&", icat_codice, icat_foglia, icat_livello, icat_padre_id, icat_ordine, icat_ordine_assoluto, "&SQL_MultiLanguageFieldComplete(conn, "icat_descr_<lingua>")&", icat_tipologia_padre_base, icat_visibile, icat_albero_visibile, icat_tipologie_padre_lista) " + _
			" 								SELECT   icat_id, "&SQL_MultiLanguageFieldComplete(conn, "icat_nome_<lingua>")&", icat_codice, icat_foglia, icat_livello, icat_padre_id, icat_ordine, icat_ordine_assoluto, "&SQL_MultiLanguageFieldComplete(conn, "icat_descr_<lingua>")&", icat_tipologia_padre_base, icat_visibile, icat_albero_visibile, icat_tipologie_padre_lista " + _
			" FROM itb_indirizzario_categorie ;"
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT tb_indirizzario_categorie OFF ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			SQL_RemoveForeignKey(conn, "tb_indirizzario", "cnt_categoria_id", "itb_indirizzario_categorie", false, "") + ";" + _
			SQL_RemoveForeignKey(conn, "irel_categ_ctech", "rcc_categoria_id", "itb_indirizzario_categorie", false, "") + ";" + _
			" DROP TABLE itb_indirizzario_categorie ;"
			
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT rel_categ_ctech ON ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			" INSERT INTO rel_categ_ctech (rcc_id, rcc_ctech_id, rcc_ordine, rcc_categoria_id) " + _
			" 					  SELECT   rcc_id, rcc_ctech_id, rcc_ordine, rcc_categoria_id " + _
			" FROM irel_categ_ctech ;"
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT rel_categ_ctech OFF ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			" DROP TABLE irel_categ_ctech ;" 
			
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT rel_cnt_ctech ON ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			" INSERT INTO rel_cnt_ctech (ric_id, ric_cnt_id, ric_ctech_id, "&SQL_MultiLanguageFieldComplete(conn, "ric_valore_<lingua>")&") " + _
			" 					  SELECT ric_id, ric_cnt_id, ric_ctech_id, "&SQL_MultiLanguageFieldComplete(conn, "ric_valore_<lingua>")&" " + _
			" FROM irel_cnt_ctech ;"
			if DB_Type(conn) = DB_SQL then
				Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + " SET IDENTITY_INSERT rel_cnt_ctech OFF ;"
			end if
			Aggiornamento__FRAMEWORK_CORE__223 = Aggiornamento__FRAMEWORK_CORE__223 + _
			" DROP TABLE irel_cnt_ctech ;"
end function
'*******************************************************************************************





'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 224
'...........................................................................................
' Giacomo 30/11/2012
'...........................................................................................
' Aggiungo parametro per attivazione sezione Attività del NextCom
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__224(conn)
    Aggiornamento__FRAMEWORK_CORE__224 = _
				" SELECT * FROM AA_VERSIONE "
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__224(conn)
	if cIntero(getValueList(conn, NULL, "SELECT id_sito FROM tb_siti WHERE id_sito=" & NEXTCOM))>0 then
		CALL AddParametroSito(conn, "NEXTCOM_ATTIVA_GESTIONE_ATTIVITA", _
							0, _
							"Attiva la sezione per la gestione delle attivita' nelle anagrafiche.", _
							"", _
							adBoolean, _
							0, _
							"", _
							1, _
							1, _
							NEXTCOM, _
							null, null, null, null, null)
	end if
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 225
'...........................................................................................
' Giacomo 30/11/2012
'...........................................................................................
' Creazione tabelle per gestione attivita' del Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__225(conn)
    Aggiornamento__FRAMEWORK_CORE__225 = _
			" ALTER TABLE tb_indirizzario_macchine ADD ima_scadenza_data smalldatetime NULL; " + _
			" " + _
			" CREATE TABLE " + SQL_Dbo(conn) + "tb_indirizzario_attivita (" + _
			"	ina_id  " & SQL_PrimaryKey(conn, "tb_indirizzario_attivita") + ", " + _
			"	ina_anagrafica_id INTEGER NULL, " + _
			" 	ina_note " + SQL_CharField(Conn, 0) + " NULL, " + _
			" 	ina_tipo_id int NOT NULL, " + _
			"	ina_data_ricontatto smalldatetime NULL, " + _
			"	ina_data_appuntamento smalldatetime NULL, " + _
			" 	ina_da_richiamare bit NULL, " + _
			" 	ina_preso_appuntamento bit NULL, " + _
			" 	ina_non_interessati bit NULL, " + _
			" 	ina_non_raggiungibili bit NULL, " + _
			"	ina_insData smalldatetime NOT NULL, " + _
			"	ina_insAdmin_id int NOT NULL, " + _
			"	ina_modData smalldatetime NULL, " + _
			"	ina_modAdmin_id int NULL); " + _
			SQL_AddForeignKey(conn, "tb_indirizzario_attivita", "ina_anagrafica_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + _
			" " + _
			" CREATE TABLE " + SQL_Dbo(conn) + "tb_indirizzario_attivita_tipi (" + _
			"	iat_id " & SQL_PrimaryKey(conn, "tb_indirizzario_attivita_tipi") + ", " + _
			" 	iat_ordine int NULL, " + _
			"	iat_icona " + SQL_CharField(Conn, 255) + " NULL, " + _
			"	iat_nome " + SQL_CharField(Conn, 255) + " NULL); " + _
			SQL_AddForeignKey(conn, "tb_indirizzario_attivita", "ina_tipo_id", "tb_indirizzario_attivita_tipi", "iat_id", false, "") + _
			" INSERT INTO tb_indirizzario_attivita_tipi(iat_nome, iat_ordine, iat_icona) VALUES('Telefonata', 10, '/amministrazione/grafica/mobile_icon.gif'); " + _
			" INSERT INTO tb_indirizzario_attivita_tipi(iat_nome, iat_ordine, iat_icona) VALUES('Visita', 20, '/amministrazione/grafica/appuntamento.png'); " + _
			" INSERT INTO tb_indirizzario_attivita_tipi(iat_nome, iat_ordine, iat_icona) VALUES('E-mail', 30, '/amministrazione/grafica/icona_email.gif'); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 226
'...........................................................................................
' Giacomo 05/12/2012
'...........................................................................................
' Creazione tabelle per gestione campagne del Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__226(conn)
    Aggiornamento__FRAMEWORK_CORE__226 = _
			" ALTER TABLE tb_indirizzario_attivita ADD ina_campagna_conclusa_id int NULL; " + _
			" " + _
			" CREATE TABLE " + SQL_Dbo(conn) + "tb_indirizzario_campagne (" + _
			"	inc_id  " & SQL_PrimaryKey(conn, "tb_indirizzario_campagne") + ", " + _
			" 	inc_nome " + SQL_CharField(Conn, 255) + " NULL, " + _
			" 	inc_note " + SQL_CharField(Conn, 0) + " NULL, " + _
			"	inc_insData smalldatetime NULL, " + _
			"	inc_insAdmin_id int NULL , " + _
			"	inc_modData smalldatetime NULL , " + _
			"	inc_modAdmin_id int NULL); " + _
			" " + _
			" CREATE TABLE " + SQL_Dbo(conn) + "rel_cnt_campagne(" + _
			"	rcc_id " & SQL_PrimaryKey(conn, "rel_cnt_campagne") + ", " + _
			"	rcc_data_conclusione smalldatetime NULL, " + _
			"	rcc_cnt_id int NULL, " + _
			"	rcc_campagna_id int NULL); " + _
			" " + _
			SQL_AddForeignKey(conn, "rel_cnt_campagne", "rcc_cnt_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + _
			SQL_AddForeignKey(conn, "rel_cnt_campagne", "rcc_campagna_id", "tb_indirizzario_campagne", "inc_id", false, "") + _
			SQL_AddForeignKey(conn, "tb_indirizzario_attivita", "ina_campagna_conclusa_id", "tb_indirizzario_campagne", "inc_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 227
'...........................................................................................
' Giacomo 14/12/2012
'...........................................................................................
' Aggiunta campi a tb_indirizzario_macchine
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__227(conn)
    Aggiornamento__FRAMEWORK_CORE__227 = _
			" ALTER TABLE tb_indirizzario_macchine ADD ima_stato_trattativa bit NULL; " + _
			" ALTER TABLE tb_indirizzario_macchine ADD ima_esito_trattativa int NULL; " + _
			" ALTER TABLE tb_indirizzario_macchine ADD ima_chiusura_trattativa_data smalldatetime NULL; "			
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 228
'...........................................................................................
' Nicola, 03/01/2013
'...........................................................................................
' aggiunta campi a struttura redirect per registrazione html e modifica relazione con indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__228(conn)
    Aggiornamento__FRAMEWORK_CORE__228 = _
			SQL_RemoveForeignKey(conn, "rel_index_url_redirect", "riu_idx_id", "tb_contents_index", true, "") + _
			" ALTER TABLE rel_index_url_redirect ADD riu_html_file " + SQL_CharField(Conn, 255) + " NULL ; " + _
			" ALTER TABLE rel_index_url_redirect ADD riu_html_data smalldatetime NULL ;" + _
			" ALTER TABLE rel_index_url_redirect ALTER COLUMN riu_idx_id INT NULL ; " + _
			SQL_AddForeignKey(conn, "rel_index_url_redirect", "riu_idx_id", "tb_contents_index", "idx_id", false, "")
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__228 = Aggiornamento__FRAMEWORK_CORE__228 + _
			DropObject(conn, "tb_contents_index_delete", "TRIGGER") + vbcRLF + _
			" CREATE TRIGGER [dbo].[tb_contents_index_delete] " + vbcRLF + _
			" ON  tb_contents_index " + vbcRLF + _
			" AFTER DELETE " + vbcRLF + _
			" AS " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			" DECLARE @idx_id_deleted int " + vbcRLF + _
			" DECLARE @is_principale bit " + vbcRLF + _
			" DECLARE @idx_content_deleted int " + vbcRLF + _
			" -- Creo un cursore per delete multipli da utilizzare come recordset sulle righe eliminate " + vbcRLF + _
			" DECLARE rs CURSOR local FAST_FORWARD FOR SELECT idx_id,idx_principale,idx_content_id FROM deleted " + vbcRLF + _
			" OPEN rs " + vbcRLF + _
			" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
			" WHILE @@FETCH_STATUS = 0 " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			" --SELECT @idx_id_deleted=idx_id,@is_principale=idx_principale,@idx_content_deleted=idx_content_id FROM deleted " + vbcRLF + _
			" IF @is_principale=0 " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			"	-- Recupero l'idx_id del nodo principale " + vbcRLF + _
			"	DECLARE @idx_principale int		" + vbcRLF + _
			"	SELECT TOP 1 @idx_principale=idx_id FROM tb_contents_index WHERE idx_content_id=@idx_content_deleted ORDER BY idx_principale DESC "	+ vbcRLF + _		 
			"	-- Associo tutti gli url dello storico al nodo principale " + vbcRLF + _
			"	IF ISNULL(@idx_principale,0) > 0 " + vbcRLF + _
			"		BEGIN " + vbcRLF + _
			"			UPDATE rel_index_url_redirect SET riu_idx_id=@idx_principale, riu_modData=GETDATE() WHERE riu_idx_id=@idx_id_deleted " + vbcRLF + _
			"			-- Recupero l'id admin da inserire nello storico " + vbcRLF + _
			"			DECLARE @id_admin int " + vbcRLF + _
			"			SELECT TOP 1 @id_admin=riu_insAdmin_id FROM rel_index_url_redirect " + vbcRLF + _
			"			-- Recupero gli url attivi per ogni lingua e se diversi da "" gli inserisco nello storico " + vbcRLF + _
			"			DECLARE @url_attivo VARCHAR(500) " + vbcRLF + _
			"			-- ITALIANO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_it,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'it',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- INGLESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_en,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'en',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- FRANCESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_fr,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'fr',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- TEDESCO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_de,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'de',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- SPAGNOLO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_es,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'es',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END	"	 + vbcRLF + _
			"			-- RUSSO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_ru,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'ru',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- CINESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_cn,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'cn',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- PORTOGHESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_pt,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect VALUES (@idx_principale,@url_attivo,'pt',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"		END " + vbcRLF + _
			" END " + vbcRLF + _
			" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
			" END " + vbcRLF + _
			" DEALLOCATE rs " + vbcRLF + _
			" END " 
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 229
'...........................................................................................
'	Nicola, 25/02/2013
'...........................................................................................
'   aggiorna vista sitemap per aggiunta link esterni che puntano all'interno del sito (ES: link utili che puntano a pagine html)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__229(conn)
	dim lingua, lingue
	if DB_Type(conn) = DB_ACCESS then
		Aggiornamento__FRAMEWORK_CORE__229 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf + _
			"    SELECT DISTINCT " + vbCrLf
			for each lingua in Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_SPAGNOLO, LINGUA_TEDESCO, LINGUA_FRANCESE)
				Aggiornamento__FRAMEWORK_CORE__229 = Aggiornamento__FRAMEWORK_CORE__229 + _
					IIF(lingua <> LINGUA_ITALIANO, ", ", "") + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " , _
							" URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " " + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri_" + lingua + vbCrLF + _
					IIF(lingua <> LINGUA_ITALIANO, ", tb_webs.lingua_" + lingua, "") + vbCrLf
			next
			Aggiornamento__FRAMEWORK_CORE__229 = Aggiornamento__FRAMEWORK_CORE__229 + _
				"        , idx_modData, id_webs, idx_livello " + vbCrLf + _
				"        FROM (v_indice_visibile INNER JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito) " + vbCrLF + _
				"             INNER JOIN tb_webs ON tb_pagineSito.id_web = tb_webs.id_webs " + vbCrLf + _
				"    	 WHERE sito_indicizzabile AND ((idx_principale OR idx_livello = 0 OR NOT idx_foglia) AND NOT riservata AND tab_per_sitemap AND indicizzabile " + _
				"				AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id)) "
	else
		Aggiornamento__FRAMEWORK_CORE__229 = _
			DropObject(conn, "v_indice_sitemap", "VIEW") + _
			" CREATE VIEW " & SQL_dbo(conn) & "v_indice_sitemap AS " + vbCrLf
			for each lingua in LINGUE_CODICI
				Aggiornamento__FRAMEWORK_CORE__229 = Aggiornamento__FRAMEWORK_CORE__229 + _
					IIF(lingua <> LINGUA_ITALIANO, " UNION " + vbCrLf, "") + _
					"    SELECT DISTINCT " + vbCrLf + _
					"       (" + SQL_IF(conn, SQL_IsTrue(conn, "URL_rewriting_attivo") & " AND (" & SQL_IfIsNull(conn, "idx_link_url_rw_" + lingua, "''") & "<>'' OR idx_link_pagina_id = id_home_page) " + vbCrLf , _
										SQL_IF(conn, "idx_link_url_rw_" + lingua + " LIKE url_base + '%' " + vbCrLf, _
											   "idx_link_url_rw_" + lingua + vbCrLf, _
											   " URL_base " & SQL_concat(conn) & "'/'" & SQL_concat(conn) & " idx_link_url_rw_" + lingua + " ") + vbCrLf, _
							" URL_base " & SQL_concat(conn) & "'/default.aspx' " & SQL_concat(conn) & " idx_link_url_" + lingua + " " + vbCrLf ) + ") AS uri, " + vbCrLF + _
					"        idx_modData, id_webs, idx_livello " + vbCrLf + _
					"        FROM (v_indice_visibile INNER JOIN tb_webs ON v_indice_visibile.idx_webs_id = tb_webs.id_webs ) " + vbCrLf + _
					"			  LEFT JOIN tb_pagineSito ON v_indice_visibile.idx_link_pagina_id = tb_pagineSito.id_pagineSito " + vbCrLF + _
					"    WHERE (" & SQL_IsTrue(conn, "sito_indicizzabile") & ") AND ((idx_link_url_" + lingua + " <> '') " & IIF(lingua <> LINGUA_ITALIANO, " AND " + SQL_IsTrue(conn, "tb_webs.lingua_" + lingua), "") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "idx_principale") & " OR idx_livello = 0 OR NOT " & SQL_IsTrue(conn, "idx_foglia") & ") " + vbCrLF + _
					"          AND (" & SQL_IsTrue(conn, "tab_per_sitemap") & ")" + vbcRLF + _
					"          AND NOT " & SQL_IsTrue(conn, "riservata") + vbCrLf + _
					"          AND (" & SQL_IsTrue(conn, "indicizzabile") & " OR (idx_link_url_IT LIKE URL_base + '%') )" + vbcRLF + _
					"		   AND (tab_name NOT LIKE 'tb_pagineSito' OR id_home_page <> co_F_key_id))"
			next
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 230
'...........................................................................................
' Giacomo 07/03/2013
'...........................................................................................
' aggiornamento per rinnovo sezione spedizione email sul Next-Com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__230(conn)
    Aggiornamento__FRAMEWORK_CORE__230 = _
			" ALTER TABLE tb_newsletters ADD nl_gestione_dinamica_contenuti bit NULL; " & _
			" UPDATE tb_newsletters SET nl_gestione_dinamica_contenuti = 1; " & _
			" ALTER TABLE log_cnt_email ADD log_messagio_letto_data smalldatetime NULL; " & _
			" ALTER TABLE tb_pagineSito ALTER COLUMN id_pagDyn_IT int NULL; " & _
			" ALTER TABLE tb_pagineSito ALTER COLUMN id_pagStage_IT int NULL; " ' in alcuni db le colonne id_pagDyn_IT e id_pagStage_IT sono NOT NULL
end function

sub AggiornamentoSpeciale__FRAMEWORK_CORE__230(conn)
	dim rs, rs_read, rs_read_2, sql, new_pagina_sito, id_template_email
	set rs = Server.CreateObject("ADODB.RecordSet")
	set rs_read = Server.CreateObject("ADODB.RecordSet")
	set rs_read_2 = Server.CreateObject("ADODB.RecordSet")
	
	sql = "SELECT * FROM tb_webs"
	rs_read.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	' per ogni sito (id_web)
	while not rs_read.eof
		id_template_email = 0
		new_pagina_sito = 0
		
		sql = " SELECT TOP 1 id_page FROM tb_pages WHERE "& SQL_IsTrue(conn, "template") & _
			  " AND " & SQL_IsTrue(conn, "semplificata") & " AND id_webs = " & rs_read("id_webs") & _
			  " ORDER BY id_page"
		rs_read_2.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
		if not rs_read_2.eof then
			id_template_email = cIntero(rs_read_2("id_page"))
		end if
		rs_read_2.close
		
		' se esiste almeno un template per email
		if id_template_email > 0 then	
			sql = "SELECT * FROM tb_pagineSito"
			rs.open sql, conn, adOpenKeySet, adLockOptimistic, adCmdText
			' inserisco una pagina sito
			rs.addNew 
			rs("id_web") = rs_read("id_webs")
			rs("archiviata") = 0
			rs("riservata") = 0
			rs("indicizzabile") = 0
			rs("ps_insData") = Now()
			rs("ps_insAdmin_id") = Session("ID_ADMIN")
			
			rs_read_2.open "SELECT * FROM tb_cnt_lingue", conn, adOpenKeySet, adLockOptimistic, adCmdText
			while not rs_read_2.eof
				rs("nome_ps_" & rs_read_2("lingua_codice")) = "Pagina per newsletter base"
				rs_read_2.moveNext
			wend
			rs_read_2.close

			rs.update
			new_pagina_sito = rs("id_paginesito")
			rs.close
			
			'creo le pagine (tb_page)
			sql = "SELECT * FROM tb_pagineSito WHERE id_paginesito = " & new_pagina_sito
			rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
			CALL Ceck_page_exists(conn, rs)
			rs.close
			'associo alle pagine il template
			sql = " UPDATE tb_pages SET id_template = " & id_template_email & _
				  " WHERE id_PaginaSito = " & new_pagina_sito
			CALL conn.execute(sql, 0, adExecuteNoRecords)
			' funzione che aggiorna il nome delle pagine nella tabella tb_pages copiandolo dalla tabella tb_pagineSito.
			CALL PaginaSitoUpdatePages(conn, new_pagina_sito)
			
		end if
		
		if new_pagina_sito > 0 then
			' inserisco la newsletter e la associo alla pagina
			sql = " INSERT INTO tb_newsletters (nl_nome_it, nl_pagina_id, nl_insAdmin_id, nl_insData, nl_lingua, nl_gestione_dinamica_contenuti) " & _
				  " VALUES ('NEWSLETTER TEMPLATE BASE - "&rs_read("nome_webs")&"', "&new_pagina_sito&", "&Session("ID_ADMIN")&", "&SQL_Date(conn, Now())&", 'it', 0)"
			CALL conn.execute(sql, 0, adExecuteNoRecords)
		end if
		
		rs_read.moveNext
	wend
	rs_read.close

	set rs = nothing
	set rs_read = nothing
	set rs_read_2 = nothing
end sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 231
'...........................................................................................
' Giacomo 13/03/2013
'...........................................................................................
' Aggiungo campo "chiave casuale" su tb_email
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__231(conn)
    Aggiornamento__FRAMEWORK_CORE__231 = _
			" ALTER TABLE tb_email ADD email_control_key " + SQL_CharField(Conn, 255) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 232
'...........................................................................................
' Giacomo 10/05/2013
'...........................................................................................
' Aggiunta campi su tabelle attività del Next-com
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__232(conn)
    Aggiornamento__FRAMEWORK_CORE__232 = _
			" ALTER TABLE tb_indirizzario_attivita ADD " & _
			" 	ina_richiamare_fatto bit NULL, " & _
			" 	ina_appuntamento_fatto bit NULL; " & _
			"UPDATE tb_indirizzario_attivita SET ina_richiamare_fatto = 1 " & _
			" WHERE ina_da_richiamare = 1 AND " & _
			"(" & SQL_CompareDateTime(conn, "ina_data_ricontatto", adCompareLessThan, DateAdd("d",-7, Now())) & ");" & _
			" UPDATE tb_indirizzario_attivita SET ina_appuntamento_fatto = 1 " & _
			" WHERE ina_preso_appuntamento = 1 AND " & _
			"(" & SQL_CompareDateTime(conn, "ina_data_appuntamento", adCompareLessThan, DateAdd("d",-7, Now())) & ")"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 233
'...........................................................................................
'	Nicola, 10/06/2013
'...........................................................................................
'   crea funzione per gestire il nome contatto da codice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__233(conn)
	if DB_Type(conn) = DB_SQL then
		Aggiornamento__FRAMEWORK_CORE__233 = _
			" CREATE FUNCTION dbo.fn_contatto_nome_completo " + vbCrLf + _
			"  (  " + vbCrLf + _
			" 	  @societa nvarchar(510),  " + vbCrLf + _
			" 	  @cognome nvarchar(255), " + vbCrLf + _
			" 	  @nome nvarchar(255) " + vbCrLf + _
			"  )  " + vbCrLf + _
			" RETURNS nvarchar(1020) " + vbCrLf + _ 
			" AS  " + vbCrLf + _
			" BEGIN  " + vbCrLf + _
			" 	  DECLARE @nomeCompleto nvarchar(1020) " + vbCrLf + _
			"     SET @nomeCompleto = LTRIM(RTRIM(IsNull(@cognome,'') + ' ' + IsNull(@nome, ''))) " + vbCrLF + _
			"     IF(IsNull(@societa,'')<>'') BEGIN " + vbCrLF + _ 
			"         IF (IsNull(@nomeCompleto,'')<>'') BEGIN " + vbCrlf + _
			"             SET @nomeCompleto = @societa + ' - ' + @nomeCompleto " + vbCrLf + _
			"         END " + vbCrlf + _
			"	      ELSE BEGIN " + vbCrLF + _
			"             SET @nomeCompleto = @societa " + vbCrLf + _
			"         END " + vbCrLF + _
			"     END " + vbCrLF + _
			"	  RETURN @nomeCompleto " + vbCrLf + _
			" END "
	end if
end function


'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 234
'...........................................................................................
' Giacomo 17/06/2013
'...........................................................................................
' Creazione tabelle per gestione sito multidominio
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__234(conn)
    Aggiornamento__FRAMEWORK_CORE__234 = _
			" CREATE TABLE " + SQL_Dbo(conn) + "tb_webs_domini (" + _
			"	dom_id  " & SQL_PrimaryKey(conn, "tb_webs_domini") + ", " + _
			" 	dom_url " + SQL_CharField(Conn, 255) + " NULL, " + _
			" 	dom_lingua " + SQL_CharField(Conn, 2) + " NULL, " + _
			" 	dom_google_analytics " + SQL_CharField(Conn, 255) + " NULL, " + _
			"	dom_web_id int NULL); " + _
			" " + _
			SQL_AddForeignKey(conn, "tb_webs_domini", "dom_web_id", "tb_webs", "id_webs", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 235
'...........................................................................................
'	Giacomo, 30/08/2013
'...........................................................................................
'   campo su tabella stati (usata nel nel contattaci)
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__235(conn)
	if TableExists(conn, "stati") then
		Aggiornamento__FRAMEWORK_CORE__235 = _
				" ALTER TABLE stati ADD " + _
				" 	isEuropa bit NULL; "
	else
		Aggiornamento__FRAMEWORK_CORE__235 = "SELECT * FROM AA_Versione"
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 236
'...........................................................................................
'	Nicola, 13/01/2014
'...........................................................................................
'   modifica gestione multidominio
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__236(conn)
	 Aggiornamento__FRAMEWORK_CORE__236 = _
			" ALTER TABLE " + SQL_Dbo(conn) + "tb_webs_domini DROP COLUMN dom_google_analytics ;" + _
			" ALTER TABLE " + SQL_Dbo(conn) + "tb_webs_domini ADD " + _
			"	dom_ordine int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 237
'...........................................................................................
'	Giacomo, 14/01/2014
'...........................................................................................
'   modifica gestione multidominio
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__237(conn)
	 Aggiornamento__FRAMEWORK_CORE__237 = _
			" ALTER TABLE " + SQL_Dbo(conn) + "tb_webs_domini ADD " + _
			"	dom_href_lang " + SQL_CharField(Conn, 50) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 238
'...........................................................................................
'	Nicola, 16/01/2014
'...........................................................................................
'   modifica gestione multidominio
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__238(conn)
	 Aggiornamento__FRAMEWORK_CORE__238 = _
			" ALTER TABLE " + SQL_Dbo(conn) + "tb_webs_domini ADD " + _
			"	dom_name " + SQL_CharField(Conn, 50) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 239
'...........................................................................................
'	Nicola, 13/02/2014
'...........................................................................................
'   aggiunta campo per gestione formati originali
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__239(conn)
	 Aggiornamento__FRAMEWORK_CORE__239 = _
			" ALTER TABLE " + SQL_Dbo(conn) + "tb_immaginiformati ADD " + _
			"	imf_salvaOriginale BIT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 240
'...........................................................................................
' Giacomo 12/05/2014
'...........................................................................................
' Aggiunta campi su tb_Indirizzario e tb_indirizzario_attivita per applicativo medici
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__240(conn)
    Aggiornamento__FRAMEWORK_CORE__240 = _
			" ALTER TABLE tb_Indirizzario ADD cnt_privato BIT NULL; " + _
			" ALTER TABLE tb_indirizzario_attivita ADD ina_descrizione_it " + SQL_CharField(Conn, 255) + " NULL; " + _
			" ALTER TABLE tb_indirizzario_attivita ADD ina_data_fine_appuntamento smalldatetime NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 241
'...........................................................................................
' Giacomo 19/05/2014
'...........................................................................................
' Aggiunta campo su tb_indirizzario_carattech
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__241(conn)
    Aggiornamento__FRAMEWORK_CORE__241 = _
			" ALTER TABLE tb_indirizzario_carattech ADD ict_righe_testo int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 242
'...........................................................................................
' Nicola 10/10/2014
'...........................................................................................
' Aggiunta campo per mantenimento informazioni su struttura url redirect
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__242(conn)
    Aggiornamento__FRAMEWORK_CORE__242 = _
			" ALTER TABLE rel_index_url_Redirect ADD " + _
		    "	riu_co_f_table_id int NULL, " + _
			"	riu_co_f_key_id int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 243
'...........................................................................................
' Nicola, 15/10/2014
'...........................................................................................
' corregge aggiornamento su trigger cancellazione indice
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__243(conn)
    Aggiornamento__FRAMEWORK_CORE__243 = DropObject(conn, "tb_contents_index_delete", "TRIGGER") + vbcRLF + _
			" CREATE TRIGGER [dbo].[tb_contents_index_delete] " + vbcRLF + _
			" ON  tb_contents_index " + vbcRLF + _
			" AFTER DELETE " + vbcRLF + _
			" AS " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			" DECLARE @idx_id_deleted int " + vbcRLF + _
			" DECLARE @is_principale bit " + vbcRLF + _
			" DECLARE @idx_content_deleted int " + vbcRLF + _
			" -- Creo un cursore per delete multipli da utilizzare come recordset sulle righe eliminate " + vbcRLF + _
			" DECLARE rs CURSOR local FAST_FORWARD FOR SELECT idx_id,idx_principale,idx_content_id FROM deleted " + vbcRLF + _
			" OPEN rs " + vbcRLF + _
			" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
			" WHILE @@FETCH_STATUS = 0 " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			" --SELECT @idx_id_deleted=idx_id,@is_principale=idx_principale,@idx_content_deleted=idx_content_id FROM deleted " + vbcRLF + _
			" IF @is_principale=0 " + vbcRLF + _
			" BEGIN " + vbcRLF + _
			"	-- Recupero l'idx_id del nodo principale " + vbcRLF + _
			"	DECLARE @idx_principale int		" + vbcRLF + _
			"	SELECT TOP 1 @idx_principale=idx_id FROM tb_contents_index WHERE idx_content_id=@idx_content_deleted ORDER BY idx_principale DESC "	+ vbcRLF + _		 
			"	-- Associo tutti gli url dello storico al nodo principale " + vbcRLF + _
			"	IF ISNULL(@idx_principale,0) > 0 " + vbcRLF + _
			"		BEGIN " + vbcRLF + _
			"			UPDATE rel_index_url_redirect SET riu_idx_id=@idx_principale, riu_modData=GETDATE() WHERE riu_idx_id=@idx_id_deleted " + vbcRLF + _
			"			-- Recupero l'id admin da inserire nello storico " + vbcRLF + _
			"			DECLARE @id_admin int " + vbcRLF + _
			"			SELECT TOP 1 @id_admin=riu_insAdmin_id FROM rel_index_url_redirect " + vbcRLF + _
			"			-- Recupero gli url attivi per ogni lingua e se diversi da "" gli inserisco nello storico " + vbcRLF + _
			"			DECLARE @url_attivo VARCHAR(500) " + vbcRLF + _
			"			-- ITALIANO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_it,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'it',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- INGLESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_en,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'en',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- FRANCESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_fr,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'fr',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- TEDESCO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_de,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'de',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- SPAGNOLO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_es,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'es',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END	"	 + vbcRLF + _
			"			-- RUSSO " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_ru,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'ru',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- CINESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_cn,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'cn',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"			-- PORTOGHESE " + vbcRLF + _
			"			SELECT @url_attivo=IsNull(idx_link_url_rw_pt,'') FROM deleted " + vbcRLF + _
			"			IF @url_attivo<>'' " + vbcRLF + _
			"			BEGIN " + vbcRLF + _
			"				INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_html_file, riu_html_data, riu_co_f_table_id, riu_co_f_key_id) VALUES (@idx_principale,@url_attivo,'pt',GETDATE(),@id_admin,GETDATE(),@id_admin, '', null, null, null) " + vbcRLF + _
			"			END " + vbcRLF + _
			"		END " + vbcRLF + _
			" END " + vbcRLF + _
			" FETCH NEXT FROM rs INTO @idx_id_deleted, @is_principale, @idx_content_deleted " + vbcRLF + _
			" END " + vbcRLF + _
			" DEALLOCATE rs " + vbcRLF + _
			" END " 
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO FRAMEWORK CORE 244
'...........................................................................................
' Luca 30/01/2015
'...........................................................................................
' Aggiunto indice per velocizzare la funzione GetUrl()
'...........................................................................................
function Aggiornamento__FRAMEWORK_CORE__244(conn)
    Aggiornamento__FRAMEWORK_CORE__244 = _
			" CREATE NONCLUSTERED INDEX [IX_tb_contents_index_geturls] ON [dbo].[tb_contents_index] " + vbcRLF + _
			" ( " + vbcRLF + _
			" 	 [idx_content_id] ASC " + vbcRLF + _
			" ) "
end function
'*******************************************************************************************
%>
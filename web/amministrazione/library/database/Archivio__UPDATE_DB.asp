<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<% 

'*******************************************************************************************
'AGGIORNAMENTO 1: creazione tabella email
'...........................................................................................
sql = " CREATE TABLE " & SQL_Dbo(Conn) & "tb_email (" + _
	  "		email_id int NOT NULL , " + _
	  "		email_text " + SQL_CharField(Conn, 0) + " NULL , " + _
	  "		email_object " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_data smalldatetime NULL , " + _
	  "		email_dipgenera int NULL , " + _
	  "		email_docs " + SQL_CharField(Conn, 0) + " NULL , " + _
	  "		email_page_ID int NULL , " + _
	  "		email_page_owned bit NOT NULL , " + _
	  "		email_in bit NOT NULL , " + _
	  "		email_MessageID " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_UIDL " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_Account int NOT NULL , " + _
	  "		email_To " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_CC " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_mime " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		email_From " + SQL_CharField(Conn, 250) + " NULL , " + _
	  "		CONSTRAINT [PK_tb_email] PRIMARY KEY  NONCLUSTERED (email_id) " + _
	  " ) "
CALL DB.Execute(sql, 1)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2 (aggiunge colonne per stato archiviazione email)
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__32(conn)
CALL DB.Execute(sql, 2)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3: allinea tabella email con le modifiche apportate con gli aggiornamenti
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__80
'Aggiornamento__FRAMEWORK_CORE__82
'Aggiornamento__FRAMEWORK_CORE__84
'...........................................................................................
sql = 	" ALTER TABLE tb_email " + SQL_AddColumn(conn) + _
	  	"		email_isBozza BIT NULL " + _
  	  	" ; " + _
		" ALTER TABLE tb_email DROP COLUMN email_page_owned; " + _
		" ALTER TABLE tb_email DROP COLUMN email_in; " + _
		" ALTER TABLE tb_email DROP COLUMN email_account; " + _
		" ALTER TABLE tb_email DROP COLUMN email_UIDL; " + _
		" ALTER TABLE tb_email DROP COLUMN email_MessageID; " + _
		" ALTER TABLE tb_email DROP COLUMN email_to; " + _
		" ALTER TABLE tb_email DROP COLUMN email_cc; " + _
		" ALTER TABLE tb_email DROP COLUMN email_from; " + _
		" UPDATE tb_email SET email_isBozza=0"
CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4: allinea tabella email con le modifiche apportate con gli aggiornamenti
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__119
'...........................................................................................
sql = 	" ALTER TABLE "& SQL_Dbo(conn) &"tb_email ADD" + vbCrLf + _
		"	email_tipi_messaggi_id INT NULL; " + vbCrLf + _
		"UPDATE tb_email SET email_tipi_messaggi_id = 1;"
CALL DB.Execute(sql, 4)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 5: allinea tabella email con le modifiche apportate con gli aggiornamenti
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__160
'...........................................................................................
sql = 	" ALTER TABLE tb_email ADD " + _
		" 	email_name_database " + SQL_CharField(Conn, 255) + " NULL ;"
		if cString(Application("DATA_ARCHIVE_ConnectionString"))<>"" then
			dim Aconn
			set Aconn = server.CreateObject("ADODB.Connection")
			Aconn.Open Application("DATA_ARCHIVE_ConnectionString"), "", "" 
			sql = sql + _ 
					" UPDATE tb_email SET email_name_database = '" & cString(Aconn.DefaultDatabase) & "'"
			Aconn.close
			set Aconn = nothing
		end if
CALL DB.Execute(sql, 5)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 6: allinea tabella email con le modifiche apportate con gli aggiornamenti
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__199
'...........................................................................................
sql = "ALTER TABLE tb_email ADD email_newsletter_tipo_id int NULL; "
CALL DB.Execute(sql, 6)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 7: crea struttura parallela per registrazione dati strutture real estate
'...........................................................................................
sql = " CREATE TABLE [dbo].[Rtb_strutture]( " + vbCrLf + _
	  " [st_ID] [int] NOT NULL, " + vbCrLf + _
	  " [st_denominazione_it] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_denominazione_en] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_denominazione_fr] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_denominazione_de] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_denominazione_es] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_it] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_en] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_fr] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_de] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_es] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_descrizione_it] [ntext] NULL, " + vbCrLf + _
	  " [st_descrizione_en] [ntext] NULL, " + vbCrLf + _
	  " [st_descrizione_fr] [ntext] NULL, " + vbCrLf + _
	  " [st_descrizione_de] [ntext] NULL, " + vbCrLf + _
	  " [st_descrizione_es] [ntext] NULL, " + vbCrLf + _
	  " [st_metratura_it] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_en] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_fr] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_de] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_es] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_ordine] [int] NULL, " + vbCrLf + _
	  " [st_home] [bit] NULL, " + vbCrLf + _
	  " [st_NEXTweb_ps_mappa_location] [int] NULL, " + vbCrLf + _
	  " [st_NEXTweb_ps_info] [int] NULL, " + vbCrLf + _
	  " [st_NEXTweb_ps_mappa_catastale] [int] NULL, " + vbCrLf + _
	  " [st_visibile] [bit] NULL, " + vbCrLf + _
	  " [st_indirizzo_mappa] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_contratto_id] [int] NULL, " + vbCrLf + _
	  " [st_categoria_id] [int] NULL, " + vbCrLf + _
	  " [st_area_id] [int] NULL, " + vbCrLf + _
	  " [st_prezzoValore_it] [money] NULL, " + vbCrLf + _
	  " [st_prezzoValore_en] [money] NULL, " + vbCrLf + _
	  " [st_prezzoValore_fr] [money] NULL, " + vbCrLf + _
	  " [st_prezzoValore_es] [money] NULL, " + vbCrLf + _
	  " [st_prezzoValore_de] [money] NULL, " + vbCrLf + _
	  " [st_agenzia_id] [int] NOT NULL, " + vbCrLf + _
	  " [st_pub_area_id] [int] NULL, " + vbCrLf + _
	  " [st_pub_contratto_id] [int] NULL, " + vbCrLf + _
	  " [st_pub_categoria_id] [int] NULL, " + vbCrLf + _
	  " [st_pub_client_id] [int] NULL, " + vbCrLf + _
	  " [st_google_maps_latitudine] [float] NULL, " + vbCrLf + _
	  " [st_google_maps_longitudine] [float] NULL, " + vbCrLf + _
	  " [st_riferimento] [nvarchar](100) NULL, " + vbCrLf + _
	  " [st_pub_visibile] [bit] NULL, " + vbCrLf + _
	  " [st_insData] [datetime] NULL, " + vbCrLf + _
	  " [st_insAdmin_id] [int] NULL, " + vbCrLf + _
	  " [st_modData] [datetime] NULL, " + vbCrLf + _
	  " [st_modAdmin_id] [int] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_it] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_en] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_fr] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_de] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_es] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_denominazione_it] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_pub_denominazione_en] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_pub_denominazione_fr] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_pub_denominazione_de] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_pub_denominazione_es] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_descrizione_ru] [ntext] NULL, " + vbCrLf + _
	  " [st_denominazione_ru] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_ru] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_ru] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzoValore_ru] [money] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_ru] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_denominazione_ru] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_descrizione_cn] [ntext] NULL, " + vbCrLf + _
	  " [st_denominazione_cn] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_cn] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_cn] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzoValore_cn] [money] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_cn] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_denominazione_cn] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_is_condominio] [bit] NULL, " + vbCrLf + _
	  " [st_condominio_id] [int] NULL, " + vbCrLf + _
	  " [st_descrizione_pt] [ntext] NULL, " + vbCrLf + _
	  " [st_denominazione_pt] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzo_pt] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metratura_pt] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_prezzoValore_pt] [money] NULL, " + vbCrLf + _
	  " [st_pub_descrizione_pt] [ntext] NULL, " + vbCrLf + _
	  " [st_pub_denominazione_pt] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_proprietario] [nvarchar](255) NULL, " + vbCrLf + _
	  " [st_metraturaValore_it] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_en] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_fr] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_de] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_es] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_cn] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_pt] [int] NULL, " + vbCrLf + _
	  " [st_metraturaValore_ru] [int] NULL, " + vbCrLf + _
	  " [st_url_it] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_en] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_fr] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_de] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_es] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_cn] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_pt] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_url_ru] [nvarchar](500) NULL, " + vbCrLf + _
	  " [st_foto_thumb] [nvarchar](255) NULL, " + vbCrLf + _
	  " CONSTRAINT [PK_Rtb_strutture] PRIMARY KEY CLUSTERED ([st_ID] DESC) " + vbCrLf + _
	  " ); " + vbCrLf + _
	  " CREATE TABLE [dbo].[Rtb_foto]( " + vbCrLf + _
	  " [fo_ID] [int] NOT NULL, " + vbCrLf + _
	  " [fo_thumb] [nvarchar](255) NULL, " + vbCrLf + _
	  " [fo_zoom] [nvarchar](255) NULL, " + vbCrLf + _
	  " [fo_didascalia_it] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_en] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_fr] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_de] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_es] [ntext] NULL, " + vbCrLf + _
	  " [fo_ordine] [int] NULL, " + vbCrLf + _
	  " [fo_struttura_id] [int] NULL, " + vbCrLf + _
	  " [fo_visibile] [bit] NULL, " + vbCrLf + _
	  " [fo_numero] [int] NULL, " + vbCrLf + _
	  " [fo_pubblicazione] [datetime] NULL, " + vbCrLf + _
	  " [fo_tipo_id] [int] NULL, " + vbCrLf + _
	  " [fo_didascalia_ru] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_cn] [ntext] NULL, " + vbCrLf + _
	  " [fo_didascalia_pt] [ntext] NULL, " + vbCrLf + _
	  " CONSTRAINT [PK_Rtb_foto] PRIMARY KEY CLUSTERED ([fo_ID] ASC) " + vbCrLf + _
	  " ); " + vbCrLf + _
	  " CREATE TABLE [dbo].[Rtb_richieste_info]( " + vbCrLf + _
	  " [ri_ID] [int] NOT NULL, " + vbCrLf + _
	  " [ri_prezzo] [money] NULL, " + vbCrLf + _
	  " [ri_richiesta] [ntext] NULL, " + vbCrLf + _
	  " [ri_NEXTcom_ID] [int] NULL, " + vbCrLf + _
	  " [ri_data] [datetime] NULL, " + vbCrLf + _
	  " [ri_struttura_id] [int] NULL, " + vbCrLf + _
	  " [ri_codice] [nvarchar](255) NULL, " + vbCrLf + _
	  " [ri_pub_id] [int] NULL, " + vbCrLf + _
	  " CONSTRAINT [PK_Rtb_richieste_info] PRIMARY KEY CLUSTERED ([ri_ID] ASC) " + vbCrLf + _
	  " ); " + vbCrLf + _
	  " CREATE TABLE [dbo].[Rrel_descrittori_realestate]( " + vbCrLf + _
	  " [rdi_ID] [int] NOT NULL, " + vbCrLf + _
	  " [rdi_descrittore_id] [int] NULL, " + vbCrLf + _
	  " [rdi_valore_it] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_valore_en] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_valore_fr] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_valore_de] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_valore_es] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_st_id] [int] NULL, " + vbCrLf + _
	  " [rdi_memo_it] [ntext] NULL, " + vbCrLf + _
	  " [rdi_memo_en] [ntext] NULL, " + vbCrLf + _
	  " [rdi_memo_fr] [ntext] NULL, " + vbCrLf + _
	  " [rdi_memo_de] [ntext] NULL, " + vbCrLf + _
	  " [rdi_memo_es] [ntext] NULL, " + vbCrLf + _
	  " [rdi_valore_ru] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_memo_ru] [ntext] NULL, " + vbCrLf + _
	  " [rdi_valore_cn] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_memo_cn] [ntext] NULL, " + vbCrLf + _
	  " [rdi_valore_pt] [nvarchar](255) NULL, " + vbCrLf + _
	  " [rdi_memo_pt] [ntext] NULL, " + vbCrLf + _
	  " CONSTRAINT [PK_Rrel_descrittori_realestate] PRIMARY KEY CLUSTERED ([rdi_ID] ASC) " + vbCrLf + _
	  " ); "
CALL DB.Execute(sql, 7)
'*******************************************************************************************


'*******************************************************************************************
'Aggiungo campo "chiave casuale" su tb_email
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__231
'...........................................................................................
sql = Aggiornamento__FRAMEWORK_CORE__231(conn)
CALL DB.Execute(sql, 8)
'*******************************************************************************************


'*******************************************************************************************
'Aggiungo tabella log_framework per archiviazione
'del framework core:
'Aggiornamento__FRAMEWORK_CORE__185
'...........................................................................................
sql = "CREATE TABLE " & SQL_dbo(conn) & "log_framework("+_
	"log_id int, "+_
	"log_table_nome NVARCHAR(50) NULL, "+_
	"log_record_id int, "+_
	"log_codice NVARCHAR(50) NULL, "+_
	"log_descrizione NVARCHAR(255) NULL, "+_
	"log_data smalldatetime, "+_
	"log_admin_id int, "+_
	"log_admin_name NVARCHAR(255) NULL, "+_
	"log_user_id int, "+_
	"log_user_name NVARCHAR(255) NULL, "+_
	"log_http_request NTEXT NULL, "+_
	"log_application_id int, "+_
	"log_application_name NVARCHAR(255) NULL);"
CALL DB.Execute(sql, 9)
'*******************************************************************************************


'*******************************************************************************************
'Aggiungo tabelle per ordini clienti - storicizzazione ordini vecchi B2B
'...........................................................................................
sql = ReadFileContent(Server.MapPath("subscripts/Aggiornamento_ARCHIVIO_10.sql"))
CALL DB.Execute(sql, 10)
'******************************************************************************************* 

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
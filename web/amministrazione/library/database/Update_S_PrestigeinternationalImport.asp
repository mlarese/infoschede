<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<%


'*******************************************************************************************
'AGGIORNAMENTO 1
'...........................................................................................
'aggiunge colonne per la gestione degli intervalli di ripulitura della cache.
'...........................................................................................
sql = " CREATE TABLE dbo.tb_trace( " + vbCrLf + _
	  "	trace_id "& SQL_PrimaryKey(conn, "tb_trace") + ", " + vbCrLf + _
	  "		trace_date smalldatetime NULL, " + vbCrLf + _
	  "		trace_operation nvarchar(4000) NULL, " + vbCrLf + _
	  " ) ; "
CALL DB.Execute(sql, 1)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'aggiunge colonne per la registrazione dei file scaricati
'...........................................................................................
sql = " ALTER TABLE dbo.tb_trace ADD " + vbCrLf + _
	  "		trace_file ntext NULL ; "
CALL DB.Execute(sql, 2)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 3
'...........................................................................................
'	Nicola, 08/11/2011
'...........................................................................................
'aggiunge tabelle per gestione import di riepilogo per verifica disponibiiltà immobili
'popola già la tabella con i dati del primo import.
'...........................................................................................
sql = " CREATE TABLE dbo.tb_import( " + vbCrLf + _
	  "	sysid "& SQL_PrimaryKeyInt(conn, "tb_import") + vbCrLf + _
	  " ) ; " + vbCrLf + _
	  " CREATE TABLE dbo.tb_import_new( " + vbCrLf + _
	  "	sysid "& SQL_PrimaryKeyInt(conn, "tb_import_new") + vbCrLf + _
	  " ) ; " & vbCrLf + _
	  " INSERT INTO tb_import (sysid) " + vbcRlf + _
	  " SELECT st_pub_client_id FROM PrestigeInternational.dbo.rtb_strutture WHERE IsNull(st_pub_client_id,0)>0 AND st_agenzia_id = 2335 AND IsNull(st_visibile,0)=1 "
CALL DB.Execute(sql, 3)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 4
'...........................................................................................
'	Nicola, 09/11/2011
'...........................................................................................
'cambia tipologia chiave primaria, aggiungendo colonna apposita per possibili sovrapposizioni
'...........................................................................................
sql = " ALTER TABLE dbo.tb_import_new 	DROP CONSTRAINT PK_tb_import_new ;" + _
	  " ALTER TABLE dbo.tb_import 		DROP CONSTRAINT PK_tb_import ;" + _
	  " ALTER TABLE dbo.tb_import 		ADD import_id " + SQL_PrimaryKey(conn, "tb_import")  + ";" + _
	  " ALTER TABLE dbo.tb_import_new 	ADD import_id " + SQL_PrimaryKey(conn, "tb_import_new")  + ";"
CALL DB.Execute(sql, 4)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 5
'...........................................................................................
'	Nicola, 09/11/2011
'...........................................................................................
'rimuove tabella non utilizzata
'...........................................................................................
sql = " DROP TABLE dbo.tb_import_new"
CALL DB.Execute(sql, 5)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 6
'...........................................................................................
'	Nicola, 05/10/2012
'...........................................................................................
'aggiunge tabelle per gestione import di riepilogo per verifica disponibiiltà immobili di New york
'popola già la tabella con i dati del primo import.
'...........................................................................................
sql = " CREATE TABLE dbo.tb_import_newyork( " + vbCrLf + _
	  "	MLS_ACCT "& SQL_PrimaryKeyInt(conn, "tb_import_newyork") + vbCrLf + _
	  " ) ; " + vbCrLf
CALL DB.Execute(sql, 6)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 7
'...........................................................................................
'	Nicola, 06/10/2012
'...........................................................................................
'rimuove tabella e ricrea
'...........................................................................................
sql = " DROP TABLE tb_import_newyork ; " + _
	  " CREATE TABLE dbo.tb_import_newyork( " + vbCrLf + _
	  " import_id "& SQL_PrimaryKey(conn, "tb_import_newyork") + vbCrLf + _
	  " , MLS_ACCT int not null " + _
	  " )"
CALL DB.Execute(sql, 7)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 8
'...........................................................................................
'	Nicola, 24/09/2015
'...........................................................................................
'rimuove tabella e ricrea
'...........................................................................................
sql = " DROP TABLE tb_import_newyork ; " + _
	  " CREATE TABLE dbo.tb_import_newyork( " + vbCrLf + _
	  " import_id "& SQL_PrimaryKey(conn, "tb_import_newyork") + vbCrLf + _
	  " , listing_id nvarchar(50) null " + _
	  " )"
CALL DB.Execute(sql, 8)
'*******************************************************************************************


%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
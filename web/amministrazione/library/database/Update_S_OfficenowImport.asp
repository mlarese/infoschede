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
sql = " ALTER TABLE tb_configurazione ADD " + _
      "     conf_cache_intervallo INT NULL, " + _
      "     conf_cache_LastClear SMALLDATETIME ; " + _
      " DELETE FROM tb_is_dispo_import; " + _
      " DELETE FROM tb_is_prezzi_import; " + _
      " DELETE FROM tb_sp_import; " + _
      " UPDATE tb_configurazione SET conf_cache_intervallo = 7, conf_cache_LastClear = GETDATE(); "
CALL DB.Execute(sql, 1)
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO 2
'...........................................................................................
'aggiunge colonne per la gestione degli intervalli di ripulitura della cache.
'...........................................................................................
sql = " ALTER TABLE tb_configurazione ADD " + _
      "     conf_cache_NextClear SMALLDATETIME ; " + _
      " UPDATE tb_configurazione SET conf_cache_NextClear = CONVERT(DATETIME, CAST(Year(GETDATE() + 7) AS nvarchar(4)) + '-' + CAST(Month(GETDATE() + 7) AS nvarchar(2)) + '-' + CAST(Day(GETDATE() + 7) AS nvarchar(2)) + ' 00:00:00', 102); "
CALL DB.Execute(sql, 2)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
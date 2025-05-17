<!--#INCLUDE FILE="Update__FileHeader.asp" -->
<% '........................................................................................... %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ToolsDescrittori.asp" -->
<%

'*******************************************************************************************
'AGGIORNAMENTO 1
'...........................................................................................
'aggiunge i campi per la gestione delle pagine per le attivita e le pubblicazioni ISE
'...........................................................................................
sql = " ALTER TABLE tb_attivita ADD COLUMN att_paginaSito_id INT NULL;" + _
	  " ALTER TABLE tb_articoli ADD COLUMN art_paginaSito_id INT NULL"
CALL DB.Execute(sql, 1)
'*******************************************************************************************

%>
<% '........................................................................................... %>
<!--#INCLUDE FILE="Update__FileFooter.asp" -->
<% '........................................................................................... %>
<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie di righe d'ordine - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA TIPOLOGIA"
dicitura.link_new = "Tabelle.asp;OrdiniRigheTipologieNew.asp"
dicitura.scrivi_con_sottosez()

dim name_session_sql
name_session_sql = "B2B_TIPORIGHE_SQL"

%>

<!--#INCLUDE FILE="Tools_OrdiniRigheTipologie.asp" -->

<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/ClassPageNavigator.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione informazioni aggiuntive per riga d'ordine - elenco"
dicitura.puls_new = "INDIETRO A TABELLE;NUOVA INFORMAZIONE"
dicitura.link_new = "Tabelle.asp;OrdiniRigheInfoNew.asp"
dicitura.scrivi_con_sottosez()  

dim name_session_sql
name_session_sql = "B2B_DetOrdInfo_SQL"

%>

<!--#INCLUDE FILE="Tools_OrdiniRigheInfo.asp" -->
<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniRigheTipologieSalva.asp")
end if

dim name_session_sql
name_session_sql = "B2B_TIPORIGHE_SQL"

%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie righe d'ordine - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "OrdiniRigheTipologie.asp"
dicitura.scrivi_con_sottosez() 

%>

<!--#INCLUDE FILE="Tools_OrdiniRigheTipologieMod.asp" -->


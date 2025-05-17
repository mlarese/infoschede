<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("VariantiSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura, name_session_sql, from_tour
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione varianti - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "Varianti.asp"
dicitura.scrivi_con_sottosez() 

name_session_sql = "B2B_VARIANTI_SQL"
from_tour = false

%>

<!--#INCLUDE FILE="Tools_VariantiMod.asp" -->

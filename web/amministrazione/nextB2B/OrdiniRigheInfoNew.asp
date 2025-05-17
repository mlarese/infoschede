<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniRigheInfoSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->

<%
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione informazioni aggiuntive per riga d'ordine - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "OrdiniRigheInfo.asp"
dicitura.scrivi_con_sottosez() 

%>

<!--#INCLUDE FILE="Tools_OrdiniRigheInfoNew.asp" -->

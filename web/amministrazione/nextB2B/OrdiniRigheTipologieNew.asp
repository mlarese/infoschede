<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<%
if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("OrdiniRigheTipologieSalva.asp")
end if
%>
<!--#INCLUDE FILE="intestazione.asp" -->
<% 	
dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione tipologie righe d'ordine - nuova"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "OrdiniRigheTipologie.asp"
dicitura.scrivi_con_sottosez()

%>

<!--#INCLUDE FILE="Tools_OrdiniRigheTipologieNew.asp" -->
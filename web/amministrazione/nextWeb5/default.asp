<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
'inizializza la sessione per l'applicativo corrente
InitSex(NEXTWEB5)

'verifica permessi di accesso
If Session("WEB_ADMIN")<>"" OR session("WEB_POWER") <> "" OR session("WEB_USER") <> "" Then
	CALL AutenticatedRedirect("Siti.asp")
else
	CALL ReturnToLogin()
End If
%>
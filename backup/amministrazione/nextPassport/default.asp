<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
'inizializza la sessione per l'applicativo corrente
InitSex(NextPassport)

'verifica permessi di accesso
If Session("PASS_ADMIN")<>"" OR Session("PASS_AMMINISTRATORI")<>"" Then
	CALL AutenticatedRedirect("Amministratori.asp")
elseIf Session("PASS_UTENTI")<>"" Then
	CALL AutenticatedRedirect("Utenti.asp")
else
	CALL ReturnToLogin()
End If
%>
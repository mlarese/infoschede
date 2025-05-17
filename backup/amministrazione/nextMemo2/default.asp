<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/ToolsDescrittori.ASP" -->
<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<%
'inizializza la sessione per l'applicativo corrente
CALL InitSex(NextMemo2)

'carica parametri del next-passport
CALL Parametri.LoadAllParams(null, null, NEXTPASSPORT)

'verifica permessi di accesso
If Session("MEMO2_ADMIN")<>"" Then
	CALL AutenticatedRedirect("Documenti.asp")
elseif Session("MEMO2_DOWNLOAD")<>"" then
	CALL AutenticatedRedirect("Download.asp")
else
	CALL ReturnToLogin()
End If
%>
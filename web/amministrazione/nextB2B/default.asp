<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
'inizializza la sessione per l'applicativo corrente
CALL InitSex(NEXTB2B)

'disabilita i tipi dei descrittori
session("DES_TIPI_DISABLE") = adIUnknown &", "& adDouble &","& adIDispatch &","& adSingle

'verifica permessi di accesso
If Session("B2B_ADMIN")<>"" Then
	CALL AutenticatedRedirect("Articoli.asp")
else
	CALL ReturnToLogin()
End If
%>
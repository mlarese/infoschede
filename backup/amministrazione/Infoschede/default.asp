<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>

<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%

'elenca parametri da controllare
CALL InitSex(INFOSCHEDE)
'CALL PreserveInitSex(NEXTB2B, true)

'carica parametri del next-com
CALL Parametri.LoadAllParams(null, null, NEXTCOM)
'carica parametri del next-b2b
CALL Parametri.LoadAllParams(null, null, NEXTB2B)

'verifica permessi di accesso
If Session("INFOSCHEDE_ADMIN")<>"" OR Session("INFOSCHEDE_CENTRO_ASSISTENZA")<>"" OR Session("INFOSCHEDE_OFFICINA")<>"" then
	CALL AutenticatedRedirect("Schede.asp")
else
	returnToLogin()
End If

%>
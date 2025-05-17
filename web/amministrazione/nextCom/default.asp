<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/INITSEX.ASP" -->
<!--#INCLUDE FILE="../library/ToolsDescrittori.asp" -->
<%
'inizializza la sessione per l'applicativo corrente
InitSex(NEXTCOM)

'verifica permessi di accesso
If Session("COM_USER")<>"" OR Session("COM_ADMIN")<>"" OR Session("COM_POWER")<>"" Then
	'disabilita i tipi dei descrittori
	session("DES_TIPI_DISABLE") = adDouble &","& adIDispatch &","& adSingle
	
	'lettura id dell'utente ed elenco rubriche visibili
	dim conn, rs, sql
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open Application("DATA_ConnectionString")
	set rs = Server.CreateObject("ADODB.RecordSet")
	
	'recupera elenco dei gruppi a cui appartiene il dipendente
	sql = "SELECT id_gruppo FROM tb_rel_dipgruppi WHERE id_impiegato =" & Session("ID_ADMIN")
	Session("DIP_GROUP") = GetValueList(conn, rs, sql)
	
	if Session("DIP_GROUP") <> "" then
		CALL AutenticatedRedirect("Contatti.asp")
	else
		Session("ERRORE") = "UTENTE NON APPARTENENTE AD UN GRUPPO DI LAVORO DELL'APPLICATIVO SELEZIONATO"
		CALL ReturnToLogin()
	end if
	
else
	CALL ReturnToLogin()
End If
%>
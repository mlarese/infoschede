<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->
<!--#INCLUDE FILE="../ClassCryptography.asp"-->
<%
dim conn, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DATA_ConnectionString"), "", ""

sql = " INSERT INTO tb_admin (admin_nome, admin_cognome, admin_email, admin_login, admin_password) " + _
	  " VALUES('Supporto tecnico', 'Combinario', 'supporto@combinario.com', 'COMBINARIO', '98E6A0FB4485D07E42FD5FEC2A66EDC4E82937E19A580B9F87') " + vbCrLF + _
	  " DECLARE @adminId int " + vbCrLF + _
	  " SELECT @adminId = id_admin FROM tb_admin WHERE admin_login LIKE 'COMBINARIO' " + vbCrLf + _
	  " INSERT INTO rel_admin_sito (admin_id, sito_id, rel_as_permesso) " + vbCrLf + _
	  " VALUES (@adminId, 1, 1) "
CALL conn.execute(sql, ,adExecuteNoRecords)

'invio e-mail all'amministratore (sviluppo@combinario.com): Manutenzione
CALL SendEmailSupportEX(Request.ServerVariables("SERVER_NAME")&" - Ricreato utente Combinario", GetRawHttp()) 
	

response.write "OK"

%>
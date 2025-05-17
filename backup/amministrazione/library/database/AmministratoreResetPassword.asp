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

sql = "UPDATE tb_admin SET admin_password = '98E6A0FB4485D07E42FD5FEC2A66EDC4E82937E19A580B9F87' WHERE admin_login LIKE 'combinario' OR admin_login LIKE 'NEXTAIM'"
CALL conn.execute(sql, ,adExecuteNoRecords)

'invio e-mail all'amministratore (sviluppo@combinario.com): Manutenzione
CALL SendEmailSupportEX(Request.ServerVariables("SERVER_NAME")&" - Resettata password per l''utente NEXT-AIM o Combinario", GetRawHttp()) 
	

response.write "OK"

%>
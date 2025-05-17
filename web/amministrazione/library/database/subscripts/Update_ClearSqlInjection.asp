<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% Server.ScriptTimeout = 1073741824 %>
<!--#INCLUDE FILE="../Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../../tools.asp" -->
<%

dim chiavi_da_troncare 
chiavi_da_troncare = Array("<SPAN STYLE=""display:none"">",_
						   "<SPAN STYLE=""display: none"">",_
						   "<span style='display:none'>",_
						   "<span style='display: none'>" _
						   )

dim conn, rs, sql, i, TRUNCATE_HTML_KEY
set conn = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")
conn.open Application("L_conn_ConnectionString")

response.write "Inizio Aggiornamento<br>"
conn.BeginTrans

'response.write server.htmlencode(GetValueList(conn, rs, "SELECT [html] FROM tb_layers WHERE id_lay=901"))

for i = lbound(chiavi_da_troncare) to ubound(chiavi_da_troncare)
	TRUNCATE_HTML_KEY = chiavi_da_troncare(i)
	sql = "SELECT * FROM tb_layers WHERE [html] like '%" & ParseSQL(TRUNCATE_HTML_KEY, adChar) & "%'"
	rs.open sql, conn, adOpenStatic, adLockOptimistic
response.write "count:" & rs.recordcount & "<br>"
	while not rs.eof
	'response.write rs("id_lay") & "<br>"
		rs("html") = left(cString(rs("html")), instr(1, cString(rs("html")), TRUNCATE_HTML_KEY, vbTextCompare)-1)
		rs.update
		
		rs.movenext
	wend

	rs.close
next

'response.write server.htmlencode(GetValueList(conn, rs, "SELECT [html] FROM tb_layers WHERE id_lay=901"))

conn.committrans
response.write "AGGIORNAMENTO ESEGUITO"

conn.close
set rs = nothing
set conn = nothing

%>
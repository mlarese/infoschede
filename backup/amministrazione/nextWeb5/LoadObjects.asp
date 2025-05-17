<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<% response.buffer = true
response.ContentType = "text/xml"

dim PageId, WebId, ObjName
PageID = cInteger(request("PAGINA"))
ObjName = cString(request("OBJ"))

dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

'recupera dati sito
sql = "SELECT id_webs FROM tb_pages WHERE id_page=" & PageId
WebId = cInteger(GetValueList(conn, rs, sql))

if ObjName = "" then
	'recupera lista dei plugin
	dim list
	sql = "SELECT * FROM tb_objects WHERE id_webs=" & WebId & " ORDER BY name_objects"
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	while not rs.eof
		list = list & rs("name_objects") & vbCrLf
		rs.movenext
	wend
	rs.close
	response.clear
	%><?xml version="1.0" encoding="UTF-8"?>
	<xml>
		<tipo nome="/objects">
			<classes><%= list %></classes>
		</tipo>
	</xml><%
else
	'recupera dati del plugin richiesto
	sql = "SELECT * FROM tb_objects WHERE name_objects LIKE '" & ObjName & "' AND id_webs=" & WebId & ""
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	response.clear
	%><?xml version="1.0" encoding="UTF-8"?>
	<xml>
		<class name="<%= rs("name_objects") %>">
			<properties><%= replace(Server.HtmlEncode(cString(rs("param_list"))), ";" & vbCrLF, "; ") %></properties>
		</class>
	</xml><%
	rs.close
end if

conn.close
set rs = nothing
set conn = nothing
%>
<% @ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% response.buffer = true %>
<%
dim layer_id
layer_id = cInteger(request("ID"))

dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

sql = "SELECT testo FROM tb_layers WHERE id_lay=" & layer_id
rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

response.clear
response.ContentType = "text/plain"
if not rs.eof then
	response.write rs("testo")
end if

rs.close

conn.close
set rs = nothing
set conn = nothing
response.end
%>
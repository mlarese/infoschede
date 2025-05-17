<%@ Language=VBScript CODEPAGE=65001%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="../library/TOOLS4Admin.ASP" -->
<% response.buffer = true
'response.ContentType = "text/xml"

dim PageId, WebId, onlyUnicode
PageID = cInteger(request("PAGINA"))

dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

'recupera dati sito
sql = "SELECT id_webs, lingua FROM tb_pages WHERE id_page=" & PageId
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdtext
WebId = rs("id_webs")
onlyUnicode = (rs("lingua") = LINGUA_CINESE)
rs.close

response.clear
dim cssO
set cssO = new CssManager

response.write cssO.GenerateEditorXml(conn, WebId, onlyUnicode)

set cssO = nothing
conn.close
set rs = nothing
set conn = nothing
%>
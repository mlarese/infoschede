<%@ Language=VBScript CODEPAGE=65001%>
<%'@ Language=VBScript CODEPAGE=1252%>
<% option explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/TOOLS.ASP" -->
<!--#INCLUDE FILE="Tools_NextWeb5.asp" -->
<% response.buffer = true %>
<%
dim page_id, web_id
page_id = cInteger(request("PAGINA"))


dim conn, rs, sql
set conn = server.createObject("ADODB.Connection")
set rs = server.createObject("ADODB.RecordSet")
conn.Open Application("DATA_ConnectionString"),"",""

sql = "SELECT id_webs FROM tb_pages WHERE id_page=" & page_id
web_id = cInteger(GetValueList(conn, rs, sql))

dim FilePath, Fso, FileStream, FileContent
Set fso = CreateObject("Scripting.FileSystemObject")

'calcola file path
FilePath = replace(request("FILE"), "*", "\")
if instrRev(FilePath, "\", 2, vbTextCompare)<1 then
	FilePath = "\" & FilePath
end if
FilePath = Application("IMAGE_PATH") & web_id & FilePath 

response.ContentType = "text/plain"

if fso.FileExists(FilePath) then
	set FileStream = fso.OpenTextFile(FilePath, 1, False, -2)
	FileContent = FileStream.ReadAll

	'ripulisce caratteri sporchi
	'FileContent = ClearString(FileContent, false)
	'pulisce Line feed
	FileContent = replace(FileContent, vbLf, "")
	
else
	FileContent = "vuoto"
end if

response.clear
response.write FileContent
response.end
%>